# 运行此程序前，请确保安装了以下依赖库：
# pip install streamlit pandas openpyxl zhipuai

import streamlit as st
import pandas as pd
import base64
from zhipuai import ZhipuAI

# ==========================================
# 1. 基础配置与全局变量
# ==========================================
# 使用您提供的智谱 AI API Key
ZHIPU_API_KEY = "8776cadbc359430cab2d635625ac65fe.3U3GnOC5VU2FfNKo"

# 设置页面基本属性
st.set_page_config(
    page_title="出差报销智能校对系统",
    page_icon="🧾",
    layout="wide"
)

# 初始化智谱 AI 客户端
try:
    client = ZhipuAI(api_key=ZHIPU_API_KEY)
except Exception as e:
    st.error(f"智谱 AI 客户端初始化失败，请检查 API Key。错误信息：{e}")

# ==========================================
# 2. 辅助函数
# ==========================================
def encode_image(uploaded_file):
    """将用户上传的图片文件转换为 Base64 编码，供 AI 视觉模型读取"""
    bytes_data = uploaded_file.getvalue()
    base64_str = base64.b64encode(bytes_data).decode('utf-8')
    return base64_str

def parse_excel(uploaded_file):
    """读取 Excel 文件并转换为 Markdown 表格文本，方便大模型理解结构"""
    try:
        # 读取第一张工作表
        df = pd.read_excel(uploaded_file)
        # 清理空数据
        df = df.dropna(how='all')
        # 转换为 Markdown 格式字符串
        return df.to_markdown(index=False)
    except Exception as e:
        return f"Excel 读取失败: {str(e)}"

def analyze_reimbursement(excel_text, image_files):
    """核心逻辑：调用智谱 GLM-4V 模型进行发票与报销单比对"""
    
    # 构建提示词 (Prompt)
    system_prompt = """
    你是一个专业、严谨的财务审核AI。你的任务是根据用户提供的【Excel报销单数据】和【发票图片】，进行严格的交叉核对。
    
    核对重点：
    1. 报销单中的【项目/费用类型】是否与发票明细内容一致。
    2. 报销单中的【金额】是否与发票上的总金额/价税合计完全一致。
    
    请仔细查看每一张发票，提取上面的金额和项目，并与表格中的每一行记录进行比对。
    
    请按照以下 Markdown 格式输出你的核对报告：
    
    ### ✅ 报销正确项目
    （列出匹配无误的报销项目及金额）
    
    ### ❌ 报销错误/存疑项目
    （列出金额不符、项目不符、或者在报销单里有但缺少对应发票的项目，并详细说明原因，例如：金额差额多少、找不到对应发票等）
    
    ### 📊 总结建议
    （给出整体的审核结论，例如：整体通过，或需打回重新填写）
    """
    
    # 构造请求消息内容
    content_list = [
        {"type": "text", "text": system_prompt},
        {"type": "text", "text": f"【Excel报销单数据】如下：\n{excel_text}"},
        {"type": "text", "text": "【发票图片】如下："}
    ]
    
    # 将图片以 base64 格式加入请求负载
    for img_file in image_files:
        base64_img = encode_image(img_file)
        content_list.append({
            "type": "image_url",
            "image_url": {
                "url": base64_img
            }
        })
        
    messages = [
        {"role": "user", "content": content_list}
    ]
    
    # 调用大模型
    try:
        response = client.chat.completions.create(
            model="glm-4v",  # 智谱的视觉大模型
            messages=messages,
            temperature=0.1,  # 降低随机性，保证财务审核的严谨性
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"调用智谱 API 失败，请检查网络或 API Key 状态。错误详情：{str(e)}"

# ==========================================
# 3. Streamlit 前端界面
# ==========================================
def main():
    st.title("🧳 旅游销售出差报销智能校对系统")
    st.markdown("欢迎使用财务自动化审核系统。请让销售人员在下方上传 **报销单(Excel)** 及对应的 **发票照片**，系统将调用 AI 自动核验。")
    st.divider()
    
    # 左右两栏布局
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1. 提交审核材料")
        # 上传 Excel
        excel_file = st.file_uploader("📂 请上传报销单表格 (支持 .xlsx, .xls)", type=['xlsx', 'xls'])
        
        # 上传发票图片（支持多选）
        image_files = st.file_uploader("🖼️ 请上传对应的发票/收据图片 (支持多选)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        
        # 开始审核按钮
        start_btn = st.button("🚀 开始智能校对", type="primary", use_container_width=True)
        
        # 预览区
        if excel_file is not None:
            with st.expander("预览报销单数据"):
                df_preview = pd.read_excel(excel_file)
                st.dataframe(df_preview)
                
        if image_files:
            with st.expander(f"预览已上传的发票 ({len(image_files)}张)"):
                # 将图片排成多列展示
                img_cols = st.columns(3)
                for i, img in enumerate(image_files):
                    img_cols[i % 3].image(img, use_container_width=True, caption=f"发票 {i+1}")

    with col2:
        st.subheader("2. AI 核对报告")
        
        if start_btn:
            if not excel_file:
                st.warning("⚠️ 请先上传报销单 Excel 文件！")
            elif not image_files:
                st.warning("⚠️ 请至少上传一张发票图片用于核对！")
            else:
                with st.spinner('🤖 智谱 AI 正在疯狂对账中，请稍候 (约需10-30秒)...'):
                    # 1. 提取 Excel 文本
                    excel_text = parse_excel(excel_file)
                    
                    # 2. 调用 AI 分析
                    result = analyze_reimbursement(excel_text, image_files)
                    
                    # 3. 展示结果
                    st.success("✅ 核对完成！")
                    st.markdown("---")
                    st.markdown(result)
        else:
            st.info("👈 请在左侧上传文件后，点击【开始智能校对】按钮。")

if __name__ == "__main__":
    main()
