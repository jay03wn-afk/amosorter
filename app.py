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
import re
from datetime import timedelta
import math

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
    try:
        if "firebase_key" in st.secrets:
            cert_dict = dict(st.secrets["firebase_key"])
            cert_dict["private_key"] = cert_dict["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(cert_dict)
        else:
            cred = credentials.Certificate('firebase-key.json')
    except Exception:
        cred = credentials.Certificate('firebase-key.json')

    firebase_admin.initialize_app(cred, {
        'storageBucket': firebaseConfig['storageBucket']
    })
db = firestore.client()

# 設定全局共享工作區 ID (達成多帳號共享進度)
SHARED_WORKSPACE = "global_shared_workspace"

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
# 3. 雲端同步與處理模組 (改為共享 WorkSpace)
# ==========================================
def upload_image_to_storage(img_bytes):
    bucket = storage.bucket()
    image_id = str(uuid.uuid4())
    blob = bucket.blob(f"workspaces/{SHARED_WORKSPACE}/images/{image_id}.png")
    blob.upload_from_string(img_bytes, content_type='image/png')
    url = blob.generate_signed_url(expiration=timedelta(days=3650))
    return url

def sync_user_meta():
    db.collection("workspaces").document(SHARED_WORKSPACE).set({
        "categories": st.session_state.categories,
        "files_meta": st.session_state.files_meta,
        "active_file_id": st.session_state.active_file_id
    }, merge=True)

def save_questions_to_firestore(new_questions):
    questions_ref = db.collection("workspaces").document(SHARED_WORKSPACE).collection("questions")
    for i in range(0, len(new_questions), 400):
        chunk = new_questions[i:i + 400]
        batch = db.batch()
        for q in chunk:
            batch.set(questions_ref.document(q["q_id"]), q)
        batch.commit()

def update_single_question_category(q_id, category, year_info, is_doubt=None):
    update_data = {"category": category, "year_info": year_info}
    if is_doubt is not None:
        update_data["is_doubt"] = is_doubt

    db.collection("workspaces").document(SHARED_WORKSPACE).collection("questions").document(q_id).update(update_data)
    
    for q in st.session_state.questions:
        if q["q_id"] == q_id:
            q["category"] = category
            q["year_info"] = year_info
            if is_doubt is not None:
                q["is_doubt"] = is_doubt
            break

def delete_file_data(file_id):
    questions_ref = db.collection("workspaces").document(SHARED_WORKSPACE).collection("questions")
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
    workspace_doc = db.collection("workspaces").document(SHARED_WORKSPACE).get()
    if workspace_doc.exists:
        data = workspace_doc.to_dict()
        st.session_state.categories = data.get("categories", ["生藥", "中藥", "法規", "實務"])
        st.session_state.files_meta = data.get("files_meta", [])
        st.session_state.active_file_id = data.get("active_file_id", None)
    
    raw_qs = [doc.to_dict() for doc in db.collection("workspaces").document(SHARED_WORKSPACE).collection("questions").stream()]
    st.session_state.questions = sorted(raw_qs, key=lambda x: (x.get("file_id", ""), x.get("order_index", 0)))

def parse_docx_with_images(file, file_id):
    doc = Document(file)
    parsed_q = []
    q_count = 0 
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
                                    img_url = upload_image_to_storage(img_bytes)
                                    q_blocks.append({"type": "image", "content": img_url})
                    if para_text.strip():
                        q_blocks.append({"type": "text", "content": para_text.strip()})
            
            text_only = "".join([b["content"] for b in q_blocks if b["type"]=="text"])
            if q_blocks and text_only not in [q.get('_raw_text', '') for q in parsed_q]:
                parsed_q.append({
                    "q_id": str(uuid.uuid4()), 
                    "file_id": file_id, 
                    "order_index": q_count,  
                    "blocks": q_blocks, 
                    "category": "未分類", 
                    "year_info": "", 
                    "_raw_text": text_only,
                    "is_doubt": False  
                })
                q_count += 1
    return parsed_q

def login_ui():
    st.title("系統登入")
    st.info("提示：目前系統設定為所有登入帳號皆共享同一份題庫與分類進度。", icon=":material/lightbulb:")
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
    st.sidebar.title("導覽列")
    PAGES = ["檔案管理", "資料夾管理", "題庫分類作業", "全站題目搜索與瀏覽", "題目匯出"]
    
    if st.session_state.current_page not in PAGES:
        if st.session_state.current_page == "題目瀏覽":
            st.session_state.current_page = "全站題目搜索與瀏覽"
        else:
            st.session_state.current_page = "檔案管理"
        
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
                        new_qs = parse_docx_with_images(uploaded_file, file_id)
                        save_questions_to_firestore(new_qs)
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
                    st.caption(f"考試資訊：{f.get('year', 114)} 年第 {f.get('session', 2)} 次")
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
                            
                            questions_ref = db.collection("workspaces").document(SHARED_WORKSPACE).collection("questions")
                            batch = db.batch()
                            count = 0
                            
                            for idx, q in enumerate(f_qs):
                                if q.get("category") != "未分類":
                                    new_year_info = f"{new_year}-{new_sess}-{idx + 1}"
                                    q["year_info"] = new_year_info 
                                    batch.update(questions_ref.document(q["q_id"]), {"year_info": new_year_info}) 
                                    count += 1
                                    
                                    if count >= 400:
                                        batch.commit()
                                        batch = db.batch()
                                        count = 0
                            if count > 0:
                                batch.commit()
                                
                            sync_user_meta()
                            st.rerun()
                            
                        st.divider()
                        st.markdown("**重置分類**")
                        reset_check = st.text_input("請輸入 `Check` 以確認重置", key=f"check_reset_{f_id}")
                        if st.button("確認重置", key=f"reset_{f_id}", use_container_width=True):
                            if reset_check == "Check":
                                for q in st.session_state.questions:
                                    if q.get("file_id") == f_id: 
                                        update_single_question_category(q["q_id"], "未分類", "", False)
                                f["current_index"] = 0
                                sync_user_meta()
                                st.success("已重置完成！")
                                st.rerun()
                            else:
                                st.error("輸入錯誤，請注意大小寫需為 Check", icon=":material/error:")
                                
                        if st.button("刪除檔案", key=f"del_{f_id}", type="primary", use_container_width=True):
                            delete_file_data(f_id)
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
                    icon_btn = ":material/folder_open:" if is_expanded else ":material/folder:"
                    
                    if st.button(f"{main_folder} (共 {total_qs_in_branch} 題)", key=f"toggle_{main_folder}", icon=icon_btn, use_container_width=True):
                        if is_expanded:
                            st.session_state.expanded_folders.remove(main_folder)
                        else:
                            st.session_state.expanded_folders.add(main_folder)
                        st.rerun()
                    
                    if is_expanded:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;- **{main_folder} (根目錄)** : `{main_qs_count}` 題")
                        for sub in sub_folders:
                            sub_name = sub.split("/")[-1]
                            st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;- {sub_name} : `{cat_counts.get(sub, 0)}` 題")
                            
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
            st.info("請先至「檔案管理」選擇要分類的檔案。", icon=":material/info:")
        else:
            active_f = next((f for f in st.session_state.files_meta if f["file_id"] == st.session_state.active_file_id), None)
            if not active_f: st.error("找不到檔案", icon=":material/error:")
            else:
                file_qs = [q for q in st.session_state.questions if q.get("file_id") == active_f["file_id"]]
                if not file_qs:
                    st.warning("無有效題目", icon=":material/warning:")
                else:
                    curr_idx = active_f.get("current_index", 0)
                    total_q = len(file_qs)
                    
                    st.markdown(f"**檔案：{active_f['name']}**")
                    st.progress((curr_idx + 1) / total_q, text=f"進度：第 {curr_idx + 1} 題 / 共 {total_q} 題")

                    current_q = file_qs[curr_idx]
                    
                    with st.container(border=True):
                        doubt_badge = " **(已標記為疑問)**" if current_q.get("is_doubt", False) else ""
                        st.caption(f"目前狀態： `{current_q.get('category', '未分類')}` {doubt_badge}")
                        for block in current_q.get('blocks', []):
                            if block['type'] == 'text': st.write(block['content'])
                            elif block['type'] == 'image': st.image(block['content'])

                    st.markdown("### 快速分類與標記")
                    if not st.session_state.categories:
                        st.info("尚無資料夾", icon=":material/info:")
                    else:
                        current_year_info = f"{active_f['year']}-{active_f['session']}-{curr_idx + 1}"
                        
                        is_doubt = st.checkbox("標記為疑問 (加入疑問區)", value=current_q.get("is_doubt", False), key=f"doubt_chk_{current_q['q_id']}")
                        if is_doubt != current_q.get("is_doubt", False):
                            update_single_question_category(current_q["q_id"], current_q.get("category", "未分類"), current_year_info, is_doubt)

                        c_sel_main, c_sel_sub, c_btn = st.columns([1.5, 1.5, 1], vertical_alignment="bottom")
                        selected_cat = hierarchical_select(
                            "分類至", 
                            st.session_state.categories, 
                            "classify", 
                            default_cat=st.session_state.last_used_folder, 
                            col_layout=(c_sel_main, c_sel_sub)
                        )
                        
                        if c_btn.button("確定分類", type="primary", use_container_width=True, icon=":material/check_circle:"):
                            update_single_question_category(current_q["q_id"], selected_cat, current_year_info, is_doubt)
                            
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

                        if st.session_state.recent_folders:
                            st.markdown("##### 最近使用的分類快捷鍵")
                            recent_cols = st.columns(5)
                            for i, r_cat in enumerate(st.session_state.recent_folders):
                                display_name = r_cat.split('/')[-1]
                                if recent_cols[i].button(display_name, key=f"recent_{r_cat}_{current_q['q_id']}", help=r_cat, use_container_width=True):
                                    update_single_question_category(current_q["q_id"], r_cat, current_year_info, is_doubt)
                                    
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

    # ---------- 頁面 4：全站題目搜索與瀏覽 ----------
    elif st.session_state.current_page == "全站題目搜索與瀏覽":
        st.header("全站題目搜索與瀏覽")
        
        with st.container(border=True):
            st.subheader("篩選條件")
            s_col1, s_col2, s_col3, s_col4 = st.columns([2, 1.5, 1.5, 1], vertical_alignment="bottom")
            search_text = s_col1.text_input("關鍵字搜尋 (針對題目內容)")
            
            # 使用母子階層設計的資料夾過濾
            main_folders = list(dict.fromkeys([c.split("/")[0] for c in st.session_state.categories]))
            main_options = ["所有分類", "未分類"] + main_folders
            selected_main = s_col2.selectbox("指定母資料夾", main_options, key="filter_main")

            if selected_main in ["所有分類", "未分類"]:
                filter_cat = selected_main
                s_col3.selectbox("指定子資料夾", ["-"], disabled=True, key="filter_sub")
            else:
                sub_options = ["(全部)"] + [c for c in st.session_state.categories if c == selected_main or c.startswith(selected_main + "/")]
                selected_sub = s_col3.selectbox("指定子資料夾", sub_options, key="filter_sub")
                filter_cat = selected_main if selected_sub == "(全部)" else selected_sub
            
            only_doubt = s_col4.checkbox("只看疑問區", value=False)
            
        filtered_qs = []
        for q in st.session_state.questions:
            if search_text and search_text.lower() not in q.get('_raw_text', '').lower():
                continue
                
            if filter_cat != "所有分類":
                if filter_cat == "未分類":
                    if q.get('category', '未分類') != "未分類":
                        continue
                else:
                    q_cat = q.get('category', '未分類')
                    if q_cat != filter_cat and not q_cat.startswith(filter_cat + "/"):
                        continue
                        
            if only_doubt and not q.get('is_doubt', False):
                continue
                
            filtered_qs.append(q)
            
        st.write(f"符合條件共 **{len(filtered_qs)}** 題")
        
        PAGE_SIZE = 15
        total_pages = math.ceil(len(filtered_qs) / PAGE_SIZE) if filtered_qs else 1
        
        if len(filtered_qs) > 0:
            page_num = st.number_input("頁碼", min_value=1, max_value=total_pages, value=1)
            start_idx = (page_num - 1) * PAGE_SIZE
            end_idx = start_idx + PAGE_SIZE
            
            for q in filtered_qs[start_idx:end_idx]:
                raw_text = q.get('_raw_text', '').strip()
                raw_text_single_line = " ".join(raw_text.splitlines())
                
                snippet_len = 30
                snippet = raw_text_single_line[:snippet_len] + "..." if len(raw_text_single_line) > snippet_len else raw_text_single_line
                    
                expander_icon = ":material/help:" if q.get("is_doubt", False) else ":material/article:"
                expander_title = f"[{q.get('year_info', '未標記')}] {snippet} - 分類: {q.get('category', '未分類')}"
                
                with st.expander(expander_title, icon=expander_icon):
                    for block in q.get('blocks', []):
                        if block['type'] == 'text': 
                            content = block['content']
                            # 關鍵字文字直接高亮標記紅色粗體
                            if search_text:
                                pattern = re.compile(re.escape(search_text), re.IGNORECASE)
                                content = pattern.sub(lambda m: f":red[**{m.group(0)}**]", content)
                            st.markdown(content)
                        elif block['type'] == 'image': st.image(block['content'])
                    
                    st.divider()
                    col_act1, col_act2, col_act3 = st.columns([1.2, 1.2, 2])
                    if col_act1.button("移回未分類", key=f"del_{q.get('q_id')}", icon=":material/delete:"):
                        update_single_question_category(q["q_id"], "未分類", q.get("year_info", ""))
                        st.rerun()
                        
                    doubt_btn_text = "取消疑問標記" if q.get("is_doubt", False) else "標記為疑問"
                    if col_act2.button(doubt_btn_text, key=f"toggle_doubt_{q.get('q_id')}", icon=":material/help_center:"):
                        new_doubt_status = not q.get("is_doubt", False)
                        update_single_question_category(q["q_id"], q.get("category", "未分類"), q.get("year_info", ""), new_doubt_status)
                        st.rerun()
                        
                    # 支援在瀏覽區直接重新分類
                    with col_act3.popover("重新分類", icon=":material/drive_file_move:"):
                        c_main_pop = st.container()
                        c_sub_pop = st.container()
                        new_cat = hierarchical_select("選擇新分類", st.session_state.categories, f"reclass_{q['q_id']}", col_layout=(c_main_pop, c_sub_pop))
                        if st.button("確定移動", key=f"btn_reclass_{q['q_id']}", icon=":material/check_circle:", type="primary", use_container_width=True):
                            update_single_question_category(q["q_id"], new_cat, q.get("year_info", ""), q.get("is_doubt", False))
                            st.rerun()

    # ---------- 頁面 5：題目匯出 ----------
    elif st.session_state.current_page == "題目匯出":
        st.header("客製化題庫匯出")
        
        with st.container(border=True):
            st.subheader("匯出範圍與過濾")
            st.write("請點選母資料夾展開後，勾選欲匯出的子分類：")
            
            export_cats = []
            main_folders = list(dict.fromkeys([c.split("/")[0] for c in st.session_state.categories]))
            
            cols = st.columns(3)
            for i, main in enumerate(main_folders):
                sub_folders = [c for c in st.session_state.categories if c == main or c.startswith(main + "/")]
                
                with cols[i % 3].expander(main, icon=":material/folder:"):
                    c_sa, c_da = st.columns(2)
                    if c_sa.button("全選", key=f"sa_{main}", use_container_width=True):
                        for sub in sub_folders:
                            st.session_state[f"export_chk_{sub}"] = True
                        st.rerun()
                    if c_da.button("清空", key=f"da_{main}", use_container_width=True):
                        for sub in sub_folders:
                            st.session_state[f"export_chk_{sub}"] = False
                        st.rerun()
                    
                    st.divider()
                    
                    for sub in sub_folders:
                        if f"export_chk_{sub}" not in st.session_state:
                            st.session_state[f"export_chk_{sub}"] = True
                            
                        is_checked = st.checkbox(sub, value=st.session_state[f"export_chk_{sub}"], key=f"dynamic_chk_{sub}")
                        st.session_state[f"export_chk_{sub}"] = is_checked
                        
                        if is_checked:
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
                    st.warning("請至少選擇一個資料夾", icon=":material/warning:")
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
if st.session_state.user is None and "remember_token" in st.query_params:
    try:
        token = st.query_params["remember_token"]
        user_info = auth.refresh(token)
        user_info["localId"] = user_info.get("user_id") 
        
        st.session_state.user = user_info
        load_from_cloud()
        st.rerun() 
        
    except Exception as e:
        st.error(f"自動登入失效，請重新登入。系統訊息: {e}", icon=":material/warning:")
        del st.query_params["remember_token"]

if st.session_state.user is None:
    login_ui()
else:
    main_app()
