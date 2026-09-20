import streamlit as st
import streamlit.components.v1 as components
import pyrebase
import firebase_admin
from firebase_admin import credentials, firestore, storage
from docx import Document
from docx.shared import Inches
import io
import zipfile
import requests
import uuid
from datetime import timedelta

# ==========================================
# 1. Firebase 初始化與設定
# ==========================================
firebaseConfig = {
    "apiKey": "AIzaSyA-A3dn0hgJm9JR2NGvijRc3AD2x3sRgng",
    "authDomain": "amooo-b2c35.firebaseapp.com",
    "databaseURL": "", 
    "projectId": "amooo-b2c35",
    "storageBucket": "amooo-b2c35.firebasestorage.app",
    "messagingSenderId": "858227626420",
    "appId": "1:858227626420:web:d924c8a857f8c2542662e1",
    "measurementId": "G-ZJZ3ZZP3VM"
}

try:
    firebase = pyrebase.initialize_app(firebaseConfig)
    auth = firebase.auth()
except:
    pass 

if not firebase_admin._apps:
    # 使用 try-except 捕捉本地端沒有 secrets.toml 的錯誤
    try:
        if "firebase_key" in st.secrets:
            # 雲端環境：將 secrets 轉為 Python 字典
            cert_dict = dict(st.secrets["firebase_key"])
            cert_dict["private_key"] = cert_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cert_dict)
        else:
            cred = credentials.Certificate('firebase-key.json')
    except Exception:
        # 本地環境：如果找不到 secrets 檔案或發生錯誤，直接讀取本地 JSON 實體檔案
        cred = credentials.Certificate('firebase-key.json')

    firebase_admin.initialize_app(cred, {
        'storageBucket': firebaseConfig['storageBucket']
    })
db = firestore.client()

# ==========================================
# 2. 網頁基本設定 & 狀態管理
# ==========================================
st.set_page_config(page_title="題庫分類系統", layout="wide")

components.html(
    """
    <script>
    if (!window.parent.document.getElementById('amo-keyboard')) {
        const script = window.parent.document.createElement('script');
        script.id = 'amo-keyboard';
        script.innerHTML = `
            document.addEventListener('keydown', function(e) {
                if (e.key === 'ArrowLeft') {
                    const btns = Array.from(document.querySelectorAll('button'));
                    const btn = btns.find(el => el.innerText.includes('上一題'));
                    if (btn) btn.click();
                }
                if (e.key === 'ArrowRight') {
                    const btns = Array.from(document.querySelectorAll('button'));
                    const btn = btns.find(el => el.innerText.includes('下一題'));
                    if (btn) btn.click();
                }
            });
        `;
        window.parent.document.head.appendChild(script);
    }
    </script>
    """, height=0, width=0
)

if "user" not in st.session_state:
    st.session_state.user = None
if "questions" not in st.session_state:
    st.session_state.questions = []
if "categories" not in st.session_state:
    st.session_state.categories = ["生藥", "中藥", "法規", "實務"]
if "files_meta" not in st.session_state:
    st.session_state.files_meta = []
if "active_file_id" not in st.session_state:
    st.session_state.active_file_id = None
if "current_page" not in st.session_state:
    st.session_state.current_page = "檔案管理"
if "expanded_folders" not in st.session_state:
    st.session_state.expanded_folders = set()
if "last_used_folder" not in st.session_state:
    st.session_state.last_used_folder = None
if "recent_folders" not in st.session_state:
    st.session_state.recent_folders = []

# ==========================================
# 3. 雲端同步與處理模組
# ==========================================
def upload_image_to_storage(uid, img_bytes):
    bucket = storage.bucket()
    image_id = str(uuid.uuid4())
    blob = bucket.blob(f"users/{uid}/images/{image_id}.png")
    blob.upload_from_string(img_bytes, content_type='image/png')
    url = blob.generate_signed_url(expiration=timedelta(days=3650))
    return url

def sync_user_meta():
    if st.session_state.user:
        uid = st.session_state.user['localId']
        db.collection("users").document(uid).set({
            "categories": st.session_state.categories,
            "files_meta": st.session_state.files_meta,
            "active_file_id": st.session_state.active_file_id
        }, merge=True)

def save_questions_to_firestore(uid, new_questions):
    questions_ref = db.collection("users").document(uid).collection("questions")
    for i in range(0, len(new_questions), 400):
        chunk = new_questions[i:i + 400]
        batch = db.batch()
        for q in chunk:
            batch.set(questions_ref.document(q["q_id"]), q)
        batch.commit()

def update_single_question_category(q_id, category, year_info):
    if st.session_state.user:
        uid = st.session_state.user['localId']
        db.collection("users").document(uid).collection("questions").document(q_id).update({
            "category": category, "year_info": year_info
        })
        for q in st.session_state.questions:
            if q["q_id"] == q_id:
                q["category"] = category
                q["year_info"] = year_info
                break

def delete_file_data(uid, file_id):
    questions_ref = db.collection("users").document(uid).collection("questions")
    q_docs = questions_ref.where("file_id", "==", file_id).stream()
    batch = db.batch()
    count = 0
    for doc in q_docs:
        batch.delete(doc.reference)
        count += 1
        if count >= 400:
            batch.commit()
            batch = db.batch()
            count = 0
    if count > 0: batch.commit()

    st.session_state.questions = [q for q in st.session_state.questions if q.get("file_id") != file_id]
    st.session_state.files_meta = [f for f in st.session_state.files_meta if f.get("file_id") != file_id]
    if st.session_state.active_file_id == file_id:
        st.session_state.active_file_id = st.session_state.files_meta[0]["file_id"] if st.session_state.files_meta else None
    sync_user_meta()

def load_from_cloud():
    uid = st.session_state.user['localId']
    user_doc = db.collection("users").document(uid).get()
    if user_doc.exists:
        data = user_doc.to_dict()
        st.session_state.categories = data.get("categories", ["生藥", "中藥", "法規", "實務"])
        st.session_state.files_meta = data.get("files_meta", [])
        st.session_state.active_file_id = data.get("active_file_id", None)
    
    raw_qs = [doc.to_dict() for doc in db.collection("users").document(uid).collection("questions").stream()]
    # 根據 file_id 與原始題號順序 (order_index) 排序，確保重新載入後順序不亂
    st.session_state.questions = sorted(raw_qs, key=lambda x: (x.get("file_id", ""), x.get("order_index", 0)))

def parse_docx_with_images(file, file_id, uid):
    doc = Document(file)
    parsed_q = []
    q_count = 0  # 紀錄題目的原始順序
    for table in doc.tables:
        for row in table.rows:
            q_blocks = []
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    para_text = ""
                    for run in paragraph.runs:
                        para_text += run.text
                        drawings = run._element.xpath('.//w:drawing')
                        for drawing in drawings:
                            for blip in drawing.xpath('.//a:blip'):
                                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                if rId:
                                    img_part = doc.part.related_parts[rId]
                                    img_bytes = img_part.blob
                                    if para_text.strip():
                                        q_blocks.append({"type": "text", "content": para_text.strip()})
                                        para_text = ""
                                    img_url = upload_image_to_storage(uid, img_bytes)
                                    q_blocks.append({"type": "image", "content": img_url})
                    if para_text.strip():
                        q_blocks.append({"type": "text", "content": para_text.strip()})
            
            text_only = "".join([b["content"] for b in q_blocks if b["type"]=="text"])
            if q_blocks and text_only not in [q.get('_raw_text', '') for q in parsed_q]:
                parsed_q.append({
                    "q_id": str(uuid.uuid4()), 
                    "file_id": file_id, 
                    "order_index": q_count,  # 儲存題目原始順序
                    "blocks": q_blocks, 
                    "category": "未分類", 
                    "year_info": "", 
                    "_raw_text": text_only 
                })
                q_count += 1
    return parsed_q

def login_ui():
    st.title("系統登入")
    email = st.text_input("帳號 (Email)")
    password = st.text_input("密碼", type="password")
    remember_me = st.checkbox("記住我的登入狀態", value=True)
    
    if st.button("安全登入", icon=":material/login:"):
        try:
            user = auth.sign_in_with_email_and_password(email, password)
            st.session_state.user = user
            if remember_me:
                st.query_params["remember_token"] = user['refreshToken']
            load_from_cloud()
            st.rerun()
        except Exception as e:
            st.error("登入失敗，請檢查帳號密碼。")

def get_year_from_info(year_info):
    if not year_info: return 0
    try: return int(str(year_info).split('-')[0])
    except: return 0

# 新增：層級化選擇器元件 (先選主層級，再選次層級)
def hierarchical_select(label, categories, key_prefix, default_cat=None, include_unclassified=False, col_layout=None):
    main_folders = list(dict.fromkeys([c.split("/")[0] for c in categories]))
    if include_unclassified:
        main_folders.insert(0, "未分類")
        
    default_main_idx = 0
    if default_cat:
        if default_cat == "未分類" and include_unclassified:
            default_main_idx = 0
        else:
            default_main = default_cat.split("/")[0]
            if default_main in main_folders:
                default_main_idx = main_folders.index(default_main)
    
    if col_layout is None:
        c1, c2 = st.columns(2)
    else:
        c1, c2 = col_layout[0], col_layout[1]
        
    selected_main = c1.selectbox(f"{label} (母資料夾)", main_folders, index=default_main_idx, key=f"{key_prefix}_main")
    
    if selected_main == "未分類":
        return "未分類"
        
    sub_folders = [c for c in categories if c == selected_main or c.startswith(selected_main + "/")]
    default_sub_idx = 0
    if default_cat and default_cat in sub_folders:
        default_sub_idx = sub_folders.index(default_cat)
        
    selected_sub = c2.selectbox(f"{label} (子資料夾與路徑)", sub_folders, index=default_sub_idx, key=f"{key_prefix}_sub")
    return selected_sub

# ==========================================
# 4. 主系統介面
# ==========================================
def main_app():
    uid = st.session_state.user['localId']
    
    # 側邊欄導覽
    st.sidebar.title("導覽列")
    PAGES = ["檔案管理", "資料夾管理", "題庫分類作業", "題目瀏覽", "題目匯出"]
    
    if st.session_state.current_page not in PAGES:
        st.session_state.current_page = "檔案管理"
        
    # 控制頁面跳轉
    selected_page = st.sidebar.radio("選擇功能", PAGES, index=PAGES.index(st.session_state.current_page))
    if selected_page != st.session_state.current_page:
        st.session_state.current_page = selected_page
        st.rerun()
        
    st.sidebar.divider()
    if st.sidebar.button("登出系統", icon=":material/logout:"):
        st.session_state.user = None
        st.query_params.clear()
        st.rerun()

    # ---------- 頁面 1：檔案管理 ----------
    if st.session_state.current_page == "檔案管理":
        st.header("檔案管理")
        with st.container(border=True):
            st.subheader("上傳新檔案", anchor=False)
            c1, c2 = st.columns(2)
            exam_year = c1.number_input("考試年份", 90, 150, 114)
            exam_session = c2.number_input("考試次數", 1, 10, 2)
            uploaded_file = st.file_uploader("選擇 Word 題庫檔", type=["docx"])
            
            if st.button("上傳並解析", icon=":material/cloud_upload:", type="primary"):
                if uploaded_file:
                    with st.spinner("正在解析與上傳..."):
                        file_id = str(uuid.uuid4())
                        new_qs = parse_docx_with_images(uploaded_file, file_id, uid)
                        save_questions_to_firestore(uid, new_qs)
                        st.session_state.questions.extend(new_qs)
                        st.session_state.files_meta.append({
                            "file_id": file_id, "name": uploaded_file.name,
                            "year": exam_year, "session": exam_session, "completed": False, "current_index": 0
                        })
                        st.session_state.active_file_id = file_id
                        sync_user_meta()
                    st.toast("上傳完成")
                    st.rerun()

        st.markdown("### 已上傳檔案")
        for f in st.session_state.files_meta:
            f_id = f["file_id"]
            f_qs = [q for q in st.session_state.questions if q.get("file_id") == f_id]
            total_qs = len(f_qs)
            categorized_qs = len([q for q in f_qs if q.get("category") != "未分類"])
            
            with st.container(border=True):
                col1, col2, col3, col4 = st.columns([1, 2.5, 1, 1], vertical_alignment="center")
                with col1:
                    is_completed = st.checkbox("標示完成", value=f.get("completed", False), key=f"chk_{f_id}")
                    if is_completed != f.get("completed", False):
                        f["completed"] = is_completed
                        sync_user_meta()
                with col2:
                    st.markdown(f"**{f['name']}**")
                    # 純文字顯示年份與次數資訊
                    st.caption(f"📅 考試資訊：{f.get('year', 114)} 年第 {f.get('session', 2)} 次")
                    
                    if total_qs > 0:
                        st.progress(categorized_qs / total_qs, text=f"進度: {categorized_qs} / {total_qs} 題")
                with col3:
                    if st.button("進入分類", key=f"btn_{f_id}", icon=":material/login:", type="primary", use_container_width=True):
                        st.session_state.active_file_id = f_id
                        sync_user_meta()
                        st.session_state.current_page = "題庫分類作業"
                        st.rerun()
                with col4:
                    with st.popover("管理", icon=":material/settings:", use_container_width=True):
                        st.markdown("**修改考試資訊**")
                        new_year = st.number_input("考試年份", value=f.get("year", 114), key=f"y_{f_id}")
                        new_sess = st.number_input("考試次數", value=f.get("session", 2), key=f"s_{f_id}")
                        
                        if new_year != f.get("year") or new_sess != f.get("session"):
                            f["year"] = new_year
                            f["session"] = new_sess
                            
                            # 連動更新該檔案底下所有「已分類」題目的 year_info 標記
                            questions_ref = db.collection("users").document(uid).collection("questions")
                            batch = db.batch()
                            count = 0
                            
                            for idx, q in enumerate(f_qs):
                                if q.get("category") != "未分類":
                                    new_year_info = f"{new_year}-{new_sess}-{idx + 1}"
                                    q["year_info"] = new_year_info  # 更新本地狀態
                                    batch.update(questions_ref.document(q["q_id"]), {"year_info": new_year_info})  # 更新雲端資料
                                    count += 1
                                    
                                    # Firestore batch 上限為 500，這裡設定 400 分批提交
                                    if count >= 400:
                                        batch.commit()
                                        batch = db.batch()
                                        count = 0
                                        
                            if count > 0:
                                batch.commit()
                                
                            sync_user_meta()
                            st.rerun()
                            
                        st.divider()
                        if st.button("重置分類", key=f"reset_{f_id}", use_container_width=True):
                            for q in st.session_state.questions:
                                if q.get("file_id") == f_id: update_single_question_category(q["q_id"], "未分類", "")
                            f["current_index"] = 0
                            sync_user_meta()
                            st.rerun()
                        if st.button("刪除檔案", key=f"del_{f_id}", type="primary", use_container_width=True):
                            delete_file_data(uid, f_id)
                            st.rerun()

    # ---------- 頁面 2：資料夾管理 ----------
    elif st.session_state.current_page == "資料夾管理":
        st.header("雲端資料夾管理")
        
        cat_counts = {cat: 0 for cat in st.session_state.categories}
        cat_counts["未分類"] = 0
        for q in st.session_state.questions:
            c = q.get('category', '未分類')
            if c in cat_counts: cat_counts[c] += 1

        folder_tree = {}
        for cat in st.session_state.categories:
            parts = cat.split("/")
            main_folder = parts[0]
            if main_folder not in folder_tree:
                folder_tree[main_folder] = []
            if len(parts) > 1:
                folder_tree[main_folder].append(cat)

        c_left, c_right = st.columns([1.5, 1])
        
        with c_left:
            st.markdown("### 瀏覽資料夾")
            with st.container(height=500, border=True):
                for main_folder, sub_folders in folder_tree.items():
                    main_qs_count = cat_counts.get(main_folder, 0)
                    total_qs_in_branch = main_qs_count + sum(cat_counts.get(sub, 0) for sub in sub_folders)
                    
                    is_expanded = main_folder in st.session_state.expanded_folders
                    icon = "📂" if is_expanded else "📁"
                    
                    if st.button(f"{icon} {main_folder} (共 {total_qs_in_branch} 題)", key=f"toggle_{main_folder}", use_container_width=True):
                        if is_expanded:
                            st.session_state.expanded_folders.remove(main_folder)
                        else:
                            st.session_state.expanded_folders.add(main_folder)
                        st.rerun()
                    
                    if is_expanded:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;📄 **{main_folder} (根目錄)** : `{main_qs_count}` 題")
                        for sub in sub_folders:
                            sub_name = sub.split("/")[-1]
                            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;📄 {sub_name} : `{cat_counts.get(sub, 0)}` 題")
                            
        with c_right:
            st.markdown("### 新增與編輯")
            with st.container(border=True):
                c_main, c_sub = st.columns(2)
                main_folders = list(dict.fromkeys([c.split("/")[0] for c in st.session_state.categories]))
                
                parent_main = c_main.selectbox("新增位置 (母資料夾)", ["建立在主目錄 (最外層)"] + main_folders, key="add_new_main")
                if parent_main == "建立在主目錄 (最外層)":
                    parent_cat = "主目錄"
                else:
                    sub_options = [c for c in st.session_state.categories if c == parent_main or c.startswith(parent_main + "/")]
                    parent_cat = c_sub.selectbox("新增位置 (延伸子資料夾)", sub_options, key="add_new_sub")

                new_cat_name = st.text_input("新資料夾名稱 (多個用逗號分隔)")
                if st.button("新增", type="primary", use_container_width=True):
                    if new_cat_name:
                        for n in [x.strip() for x in new_cat_name.split(",") if x.strip()]:
                            full_path = n if "主目錄" in parent_cat else f"{parent_cat}/{n}"
                            if full_path not in st.session_state.categories: st.session_state.categories.append(full_path)
                        sync_user_meta()
                        st.rerun()

            with st.container(border=True):
                target_cat = hierarchical_select("選擇資料夾", st.session_state.categories, "edit_del")
                new_name = st.text_input("重新命名為", value=target_cat.split("/")[-1] if target_cat else "")
                c_e1, c_e2 = st.columns(2)
                if c_e1.button("儲存名稱", use_container_width=True):
                    old_name = target_cat.split("/")[-1]
                    if new_name != old_name:
                        base_path = "/".join(target_cat.split("/")[:-1])
                        new_full_path = f"{base_path}/{new_name}" if base_path else new_name
                        idx = st.session_state.categories.index(target_cat)
                        st.session_state.categories[idx] = new_full_path
                        
                        for i, c in enumerate(st.session_state.categories):
                            if c.startswith(target_cat + "/"):
                                st.session_state.categories[i] = c.replace(target_cat, new_full_path, 1)
                        for q in st.session_state.questions:
                            if q.get('category') == target_cat:
                                update_single_question_category(q["q_id"], new_full_path, q.get("year_info", ""))
                            elif q.get('category', '').startswith(target_cat + "/"):
                                update_single_question_category(q["q_id"], q['category'].replace(target_cat, new_full_path, 1), q.get("year_info", ""))
                        sync_user_meta()
                        st.rerun()
                        
                if c_e2.button("刪除", type="primary", use_container_width=True):
                    to_del = [c for c in st.session_state.categories if c == target_cat or c.startswith(target_cat + "/")]
                    for c in to_del: st.session_state.categories.remove(c)
                    for q in st.session_state.questions:
                        if q.get('category') in to_del: update_single_question_category(q["q_id"], "未分類", "")
                    sync_user_meta()
                    st.rerun()

    # ---------- 頁面 3：題庫分類作業 ----------
    elif st.session_state.current_page == "題庫分類作業":
        st.header("題庫分類作業")
        if not st.session_state.active_file_id:
            st.info("請先至「檔案管理」選擇要分類的檔案。")
        else:
            active_f = next((f for f in st.session_state.files_meta if f["file_id"] == st.session_state.active_file_id), None)
            if not active_f: st.error("找不到檔案")
            else:
                file_qs = [q for q in st.session_state.questions if q.get("file_id") == active_f["file_id"]]
                if not file_qs:
                    st.warning("無有效題目")
                else:
                    curr_idx = active_f.get("current_index", 0)
                    total_q = len(file_qs)
                    
                    st.markdown(f"**檔案：{active_f['name']}**")
                    st.progress((curr_idx + 1) / total_q, text=f"進度：第 {curr_idx + 1} 題 / 共 {total_q} 題")

                    current_q = file_qs[curr_idx]
                    
                    with st.container(border=True):
                        st.caption(f"目前狀態： `{current_q.get('category', '未分類')}`")
                        for block in current_q.get('blocks', []):
                            if block['type'] == 'text': st.write(block['content'])
                            elif block['type'] == 'image': st.image(block['content'])

                    st.markdown("### 快速分類")
                    if not st.session_state.categories:
                        st.info("尚無資料夾")
                    else:
                        current_year_info = f"{active_f['year']}-{active_f['session']}-{curr_idx + 1}"
                        
                        # 改為層級選擇器佈局
                        c_sel_main, c_sel_sub, c_btn = st.columns([1.5, 1.5, 1], vertical_alignment="bottom")
                        selected_cat = hierarchical_select(
                            "分類至", 
                            st.session_state.categories, 
                            "classify", 
                            default_cat=st.session_state.last_used_folder, 
                            col_layout=(c_sel_main, c_sel_sub)
                        )
                        
                        if c_btn.button("確定分類", type="primary", use_container_width=True, icon=":material/check_circle:"):
                            update_single_question_category(current_q["q_id"], selected_cat, current_year_info)
                            
                            st.session_state.last_used_folder = selected_cat
                            if selected_cat in st.session_state.recent_folders:
                                st.session_state.recent_folders.remove(selected_cat)
                            st.session_state.recent_folders.insert(0, selected_cat)
                            if len(st.session_state.recent_folders) > 5:
                                st.session_state.recent_folders = st.session_state.recent_folders[:5]
                                
                            if curr_idx < total_q - 1:
                                active_f['current_index'] = curr_idx + 1
                                sync_user_meta()
                            st.rerun()

                        # 顯示最近使用的五個資料夾快捷列
                        if st.session_state.recent_folders:
                            st.markdown("##### 📌 最近使用的分類快捷鍵")
                            recent_cols = st.columns(5)
                            for i, r_cat in enumerate(st.session_state.recent_folders):
                                display_name = r_cat.split('/')[-1]
                                if recent_cols[i].button(display_name, key=f"recent_{r_cat}_{current_q['q_id']}", help=r_cat, use_container_width=True):
                                    update_single_question_category(current_q["q_id"], r_cat, current_year_info)
                                    
                                    st.session_state.last_used_folder = r_cat
                                    st.session_state.recent_folders.remove(r_cat)
                                    st.session_state.recent_folders.insert(0, r_cat)
                                    
                                    if curr_idx < total_q - 1:
                                        active_f['current_index'] = curr_idx + 1
                                        sync_user_meta()
                                    st.rerun()

                    st.divider()
                    c1, c2, c3, c4 = st.columns(4)
                    if c1.button("上一題", icon=":material/arrow_back:") and curr_idx > 0:
                        active_f['current_index'] = curr_idx - 1; sync_user_meta(); st.rerun()
                    if c2.button("下一題", icon=":material/arrow_forward:") and curr_idx < total_q - 1:
                        active_f['current_index'] = curr_idx + 1; sync_user_meta(); st.rerun()
                    jump_to = c3.number_input("跳題", 1, total_q, curr_idx + 1, label_visibility="collapsed")
                    if c4.button("跳轉", icon=":material/keyboard_tab:"):
                        active_f['current_index'] = jump_to - 1; sync_user_meta(); st.rerun()

    # ---------- 頁面 4：題目瀏覽 ----------
    elif st.session_state.current_page == "題目瀏覽":
        st.header("題目瀏覽與管理")
        
        # 改為層級選擇器
        filter_cat = hierarchical_select("預覽", st.session_state.categories, "browse", include_unclassified=True)
        filtered_qs = [q for q in st.session_state.questions if q.get('category') == filter_cat]
        
        st.write(f"共有 {len(filtered_qs)} 題")
        for q in filtered_qs:
            with st.expander(f"[{q.get('year_info', '未標記')}]"):
                for block in q.get('blocks', []):
                    if block['type'] == 'text': st.write(block['content'])
                    elif block['type'] == 'image': st.image(block['content'])
                if st.button("移回未分類", key=f"del_{q.get('q_id')}", icon=":material/delete:"):
                    update_single_question_category(q["q_id"], "未分類", "")
                    st.rerun()

    # ---------- 頁面 5：題目匯出 ----------
    elif st.session_state.current_page == "題目匯出":
        st.header("客製化題庫匯出")
        
        with st.container(border=True):
            st.subheader("匯出範圍與過濾")
            st.write("請點選母資料夾展開後，勾選欲匯出的子分類：")
            
            export_cats = []
            main_folders = list(dict.fromkeys([c.split("/")[0] for c in st.session_state.categories]))
            
            # 使用折疊面板 (Expander) 達成「點擊母資料夾才看得到子資料夾」
            cols = st.columns(3)
            for i, main in enumerate(main_folders):
                sub_folders = [c for c in st.session_state.categories if c == main or c.startswith(main + "/")]
                with cols[i % 3].expander(f"📁 {main}"):
                    for sub in sub_folders:
                        # 預設全勾選
                        if st.checkbox(sub, value=True, key=f"export_{sub}"):
                            export_cats.append(sub)

            st.divider()
            c_f1, c_f2 = st.columns(2)
            min_year = c_f1.number_input("最小年份", value=90, min_value=1)
            max_year = c_f2.number_input("最大年份", value=150, max_value=300)
            
            st.divider()
            st.subheader("匯出格式設定")
            c_opts1, c_opts2, c_opts3 = st.columns(3)
            export_format = c_opts1.radio("匯出結構", ["合併為單一 Word 檔", "依資料夾分別打包 (ZIP)"])
            include_year = c_opts2.checkbox("匯出年份題號標記 (如: [114-2-1])", value=True)
            include_images = c_opts2.checkbox("匯出圖片", value=True)
            custom_title = c_opts3.text_input("單一合併檔主標題", "客製化分類題庫")
            
            if st.button("開始產生文件", type="primary", icon=":material/download:"):
                if not export_cats:
                    st.warning("請至少選擇一個資料夾")
                else:
                    with st.spinner("正在產生文件，這可能需要一點時間..."):
                        if export_format == "合併為單一 Word 檔":
                            doc = Document()
                            doc.add_heading(custom_title, 0)
                            
                            for cat in export_cats:
                                all_cat_qs = [q for q in st.session_state.questions if q.get('category') == cat]
                                cat_qs = [q for q in all_cat_qs if min_year <= get_year_from_info(q.get('year_info', '')) <= max_year]
                                
                                if cat_qs:
                                    doc.add_heading(f'{cat.replace("/", " - ")}', 1)
                                    for q in cat_qs:
                                        if include_year: doc.add_paragraph(f"[{q.get('year_info', '')}]")
                                        for block in q.get('blocks', []):
                                            if block['type'] == 'text': doc.add_paragraph(block['content'])
                                            elif block['type'] == 'image' and include_images:
                                                try:
                                                    img_res = requests.get(block['content'], timeout=10)
                                                    if img_res.status_code == 200: doc.add_picture(io.BytesIO(img_res.content), width=Inches(4))
                                                except:
                                                    doc.add_paragraph("[圖片載入失敗]")
                                        doc.add_paragraph("---")
                            doc_buffer = io.BytesIO()
                            doc.save(doc_buffer)
                            st.download_button("下載完整題庫 (Docx)", doc_buffer.getvalue(), f"{custom_title}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for cat in export_cats:
                                    all_cat_qs = [q for q in st.session_state.questions if q.get('category') == cat]
                                    cat_qs = [q for q in all_cat_qs if min_year <= get_year_from_info(q.get('year_info', '')) <= max_year]
                                    
                                    if cat_qs:
                                        doc = Document()
                                        doc.add_heading(f'{cat.replace("/", "_")} 題庫', 0)
                                        for q in cat_qs:
                                            if include_year: doc.add_paragraph(f"[{q.get('year_info', '')}]")
                                            for block in q.get('blocks', []):
                                                if block['type'] == 'text': doc.add_paragraph(block['content'])
                                                elif block['type'] == 'image' and include_images:
                                                    try:
                                                        img_res = requests.get(block['content'], timeout=10)
                                                        if img_res.status_code == 200: doc.add_picture(io.BytesIO(img_res.content), width=Inches(4))
                                                    except: pass
                                            doc.add_paragraph("---")
                                        doc_buffer = io.BytesIO()
                                        doc.save(doc_buffer)
                                        zip_file.writestr(f"分類_{cat.replace('/', '_')}.docx", doc_buffer.getvalue())
                            st.download_button("下載壓縮檔 (ZIP)", zip_buffer.getvalue(), "分類題庫打包.zip", "application/zip")

# ==========================================
# 自動登入與畫面路由
# ==========================================

# 1. 自動登入檢查：若尚未登入但瀏覽器網址列記有 token，則進行自動認證
if st.session_state.user is None and "remember_token" in st.query_params:
    try:
        token = st.query_params["remember_token"]
        # 呼叫 Firebase API 刷新 Token
        user_info = auth.refresh(token)
        # Google API 刷新後回傳的鍵值通常是 'user_id'
        user_info["localId"] = user_info.get("user_id") 
        
        st.session_state.user = user_info
        load_from_cloud()
        
    except Exception as e:
        # 如果憑證過期或刷新失敗，不要默默清除，顯示出錯誤原因
        st.error(f"⚠️ 自動登入失效，請重新登入。系統訊息: {e}")
        del st.query_params["remember_token"]

# 2. 根據登入狀態渲染畫面
if st.session_state.user is None:
    login_ui()
else:
    main_app()
