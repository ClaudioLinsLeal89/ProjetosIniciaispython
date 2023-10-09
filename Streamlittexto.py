import streamlit as st

st.title('FLUXO DE CAIXA')
st.header("TOMADA DE DECISÃO")
st.subheader("ECONOMIAS E ESTOUROS")

st.markdown("Importância da assertividade das obrigações a pagar para o efeito ao caixa")

st.caption("este é  o captiom")

# Code

code = '''if(fome > 0)
      return "ir para geladeira"
else: 
      return "estudar Streamlit"'''
st.code(code,language='python')

st.text('este é um texto usando st.texto')

# latex https://katex.org/docs/support_table

st.latex("\begin{align}a&=b+c\\d+e&=f\end{align}")


