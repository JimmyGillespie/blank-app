import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Streamlit Test App", page_icon="✅", layout="centered")

st.title("✅ Streamlit Test App")
st.write("If you can see this page and interact with the widgets below, Streamlit is working correctly on your machine.")

st.divider()

# 1. Text input test
st.subheader("1. Text Input")
name = st.text_input("Enter your name", value="James")
if name:
    st.success(f"Hello, {name}! Text input is working.")

st.divider()

# 2. Button test
st.subheader("2. Button")
if st.button("Click me"):
    st.balloons()
    st.success("Button click works!")

st.divider()

# 3. Slider test
st.subheader("3. Slider")
number = st.slider("Pick a number", 0, 100, 50)
st.write(f"You picked: **{number}**")

st.divider()

# 4. Chart test (random data)
st.subheader("4. Chart")
chart_data = pd.DataFrame(
    np.random.randn(20, 3),
    columns=["A", "B", "C"]
)
st.line_chart(chart_data)

st.divider()

# 5. DataFrame test
st.subheader("5. Table")
st.dataframe(chart_data)

st.divider()
st.caption("All sections loaded without errors = your Streamlit setup is good to go.")
