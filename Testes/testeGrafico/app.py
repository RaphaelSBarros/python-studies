import streamlit as st

# Textos

#st.header('Meu Dashboard')
#st.sidebar.text('Meu texto')

#st.markdown('# Meu Título')
#st.caption('Minha legenda')

#pessoas = [
#    {'Nome': 'Caio', 'Idade': 22},
#    {'Nome': 'Marcos', 'Idade': 25}   
#]

#st.write('# Pessoa', pessoas)

# Exibição de Dados

import pandas as pd
import numpy as np



df = pd.DataFrame(
    np.random.rand(10, 4),
    columns=['Preço', 'Taxa de Ocupação', 'Taxa de Inadimplência', 'Pessoas por casa']
)

date = st.date_input("Pick a date")

print(date)

x = not st.checkbox('Mostrar tabela')

if not x:
    st.bar_chart(df)
else:
    print(date[0][2])