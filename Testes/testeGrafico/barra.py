import plotly.express as px

dados_x=['2018', '2019', '2020', '2021']
dados_y=[10, 20, 5, 35]

fig=px.bar(x=dados_x, y=dados_y, title='Teste Gráfico', width=800, height=500)

fig.update_yaxes(title='vertical') #muda o título do eixo
fig.update_xaxes(title='horizontal')

fig.update_traces(text=dados_y) #insere os dados dentro do gráfico

fig.show()