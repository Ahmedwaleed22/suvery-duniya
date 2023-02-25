# pip install python-docx 
# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
import pickle
import pandas as pd
from Google import Create_Service
from googleapiclient.errors import HttpError
import warnings
warnings.filterwarnings('ignore')

import os
import pandas as pd
from dash import Dash, dcc, html, Input, Output, dash_table
import plotly.express as px
import plotly.graph_objects as go
import sys

def create_service():

    CLIENT_SECRET_FILE = 'client_secret.json'
    API_NAME = 'forms'
    API_VERSION = 'v1'
    SCOPES = ["https://www.googleapis.com/auth/forms.body", "https://www.googleapis.com/auth/forms.responses.readonly"]

    try:
        service = pickle.load(open('service.pkl', 'rb'))
        service
    except:
        service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
        with open('service.pkl', 'wb' ) as f:
            pickle.dump(service, f)
            f.close()

    return service

def get_responses(form_id, service):
    
    form = service.forms().get(formId=form_id).execute()
    
    title = form['info']['title']
    
    question_dict = {}

    for item in form['items']:

        q_id = item['questionItem']['question']['questionId']

        question_dict[q_id] = item['title']
        
    resp = service.forms().responses().list(formId=form_id).execute()
    if resp:
        temp_list = []
        for response in resp['responses']:
            temp_dict = {}
            for q_id in question_dict:
                try:
                    answers = response['answers'][q_id]['textAnswers']['answers']
                    value = ', '.join([ele['value'] for ele in answers])
                    temp_dict[q_id] = value

                except KeyError:
                    temp_dict[q_id] = None
            
            temp_list.append(temp_dict)

        response_df = pd.DataFrame(temp_list)
        response_df.rename(columns=question_dict, inplace=True)
        
        return response_df
    else:
        print('No Responses')

def data_transform(df):
    
    try:
        is_numeric_dict = dict(df.apply(lambda x: x.str.replace('.', '', 1).str.isnumeric()).sum() > round(df.shape[0]/2))
    except AttributeError:
        df = df.astype('str')
        is_numeric_dict = dict(df.apply(lambda x: x.str.replace('.', '', 1).str.isnumeric()).sum() > round(df.shape[0]/2))
    
    for col in df.columns:
        if is_numeric_dict[col]:
            df[col] = df[col].astype(float)
        else:
            df[col] = df[col].astype(str)
            
    return df

def get_response_no(form_id):

    service = create_service()

    resp = service.forms().responses().list(formId=form_id).execute()
    if resp:
        return len(resp['responses'])
    else:
        return 0

def generate_table(dataframe):
    return dash_table.DataTable(
        dataframe.to_dict('records'), [{"name": i, "id": i} for i in dataframe.columns],
        style_header={
        'backgroundColor': 'rgb(30, 30, 30)',
        'color': 'white'
    },
    style_data={
        'backgroundColor': 'rgb(50, 50, 50)',
        'color': 'white'
    },
    fixed_rows={'headers': True},
    style_table={'height': 400}
)

def display_gauge(achieved, target):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = achieved,
        title = {'text': "Number of Respondents"},
        domain = {'x': [0, 1], 'y': [0, 1]},
        gauge = {'axis': {'range': [0, target]}}
    ))
    return fig

#Usage
#df1 = get_responses(form_id='1JViYrwlcSZnJuaCmrXHpwAGrqYY7es5TErX7QZPtpTY')

#df = pd.read_excel('sample_response.xlsx')

def init_dashboard(server, form_id):

    url_base_pathname = '/dashboard/' + form_id + '/live/'

    app = Dash(__name__, server=server, url_base_pathname=url_base_pathname)

    service = create_service()

    # df = get_responses(form_id=form_id, service=service)
    # df = data_transform(df)
    df = pd.read_excel('Untitled spreadsheet.xlsx')
    

    achieved = int(df.shape[0])
    target = 500

    cat_cols = list(df.select_dtypes('object').columns)
    num_cols = list(df.select_dtypes(['int', 'float']).columns)

    app.layout = html.Div([
        html.H1('SURVEY ANALYSIS', style={'text-align':'center'}),
        html.H2(children='Status', style={'text-align':'center'}),
        dcc.Graph(figure=display_gauge(achieved, target)),
        html.H2(children='DataFrame', style={'text-align':'center'}),
        generate_table(df),
        html.H2('PIE CHART', style={'text-align':'center'}),
        dcc.Graph(id="pie-chart"),
        html.P("Categorical Variable:"),
        dcc.Dropdown(id='pie_cats',
            options=cat_cols,
            value=cat_cols[0], clearable=False,
            placeholder="Select the Categorical Variable"
        ),
        html.P("Numerical variable:"),
        dcc.Dropdown(id='pie_nums',
            options=num_cols,
            value=num_cols[0] if num_cols else None, clearable=False,
            placeholder='Select the Numerical Variable' if num_cols else 'No Numerical Variables to Select'
        ),
        html.H2('BAR CHART', style={'text-align':'center'}),
        dcc.Graph(id="bar-chart"),
        html.P("Categorical Variable:"),
        dcc.Dropdown(id='bar_cats',
            options=cat_cols,
            value=cat_cols[0], clearable=False,
            placeholder="Select the Categorical Variable"
        ),
        html.P("Numerical variable:"),
        dcc.Dropdown(id='bar_nums',
            options=num_cols,
            value=num_cols[0] if num_cols else None, clearable=False,
            placeholder='Select the Numerical Variable' if num_cols else 'No Numerical Variables to Select'
        ),
        html.H2("DISTRIBUTION PLOT", style={'text-align':'center'}),
        dcc.Graph(id="dist-plot"),
        html.P("Select Distribution:"),
        dcc.RadioItems(
            id='distribution',
            options=['box', 'violin', 'rug'],
            value='box', inline=True
        ),
        html.P("Select Variable:"),
        dcc.Dropdown(id='dist_nums',
            options=num_cols,
            value=num_cols[0], clearable=False,
            placeholder="Select the Numerical Variable"
        ),
        html.P("Color By:"),
        dcc.Dropdown(id='dist_cats',
            options=cat_cols,
            value=cat_cols[0], clearable=False,
            placeholder="Select the Categorical Variable"
        )
    ])

    #---------------------------------------------------------------------------------------------------------#

    @app.callback(
        Output("pie-chart", "figure"),
        Input("pie_cats", "value"),
        Input("pie_nums", "value"))

    def generate_pie_chart(pie_cats, pie_nums):
        fig = px.pie(df, names=pie_cats, values=pie_nums if pie_nums else None, hole=.3)
        fig.update_layout(title_text=f'{pie_cats} and {pie_nums}' if pie_nums else f'{pie_cats}', title_x=0.5)
        fig['layout']['title']['font'] = dict(size=20)
        return fig


    @app.callback(
        Output("bar-chart", "figure"),
        Input("bar_cats", "value"),
        Input("bar_nums", "value"))

    def generate_bar_chart(bar_cats, bar_nums):
        fig = px.bar(df, x=bar_cats, y=bar_nums if bar_nums else None, color=bar_cats)
        fig.update_layout(title_text=f'{bar_cats} and {bar_nums}' if bar_nums else f'{bar_cats}', title_x=0.5)
        fig['layout']['title']['font'] = dict(size=20)
        return fig

    @app.callback(
        Output("dist-plot", "figure"),
        Input("distribution", "value"),
        Input("dist_nums", "value"),
        Input("dist_cats", "value"))

    def display_graph(distribution, dist_nums, dist_cats):
        fig = px.histogram(
            df, x=dist_nums, color=dist_cats if dist_cats else None,
            marginal=distribution, range_x=[-5, 60],
            hover_data=df.columns)
        fig.update_layout(title_text=f' Distribution of {dist_nums} by {dist_cats}' if dist_cats else f'Distribution of {dist_nums}', title_x=0.5)
        return fig

    return app


form_id = sys.argv[1]
print(get_response_no(form_id=form_id))