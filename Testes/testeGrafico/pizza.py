import plotly.express as px

dados_x=['2018', '2019', '2020', '2021']
dados_y=[10, 20, 5, 35]

fig = px.pie(names = dados_x, values=dados_y, width=300, height=300)

fig.show()