# Text-to-PPT

1. streamlit 界面入口，一个对话输入框 + 一个上传文件项
2. 用户输入prompt需求 + 用户上传原始word文档/txt文档
3. 解析文档，得到字符串
4. 大模型基于用户prompt拆分文档字符串，得到大纲、json格式数据
5. 根据json使用python-pptx生成pptx
6. streamlit 提供下载按钮
