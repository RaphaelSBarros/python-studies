import plotly.express as px

dados_x=['2018', '2019', '2020', '2021']
dados_y=[10, 20, 5, 35]

fig = px.line(x=dados_x, y=dados_y, title="Vendas x Ano", height=400, width=1000, line_shape='spline') # insere os dados na horizontal, vertical, atribui um título, altura, largura e o formato da linha do gráfico
fig.update_yaxes(title='Vendas', title_font_color='red', ticks='outside', tickfont_color='yellow') # atualiza o título do eixo Y para a cor vermelha, coloca traços entre o numero e o grafico e a fonte dos dados verticais para amarelo

fig.show()