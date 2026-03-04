import streamlit as st
import pandas as pd
import numpy as np
# 1. 設定網頁標題
st.title(&#39; 我的第一個 Streamlit App&#39;)
# 2. 顯示一段文字
st.write(&#39;恭喜你！你已經成功架設了你的第一個網頁應用程式。&#39;)
# 3. 互動元件：輸入名字
user_name = st.text_input(&quot;請輸入你的名字&quot;, &quot;訪客&quot;)
if user_name:
st.success(f&quot;哈囉, {user_name}！歡迎來到 Streamlit 的世界。&quot;)
# 4. 數據視覺化：隨機產生數據並畫圖
st.subheader(&#39;� 簡單的數據展示&#39;)
chart_data = pd.DataFrame(
np.random.randn(20, 3),
columns=[&#39;A&#39;, &#39;B&#39;, &#39;C&#39;]
)
# 直接繪製折線圖
st.line_chart(chart_data)
# 5. 側邊欄範例
with st.sidebar:
st.header(&quot;側邊欄設定&quot;)
st.write(&quot;這裡可以放參數設定或導覽列。&quot;)
st.button(&quot;沒用的按鈕&quot;)
