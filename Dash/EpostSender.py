from dash import Dash, html, dcc, exceptions
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
import plotly.express as px
import pandas as pd
import webbrowser
import base64
import datetime
import io
import win32com.client
import time
import datetime
from collections import Counter


class DashboardMaker:

    def __init__(self):
        self.app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
        self.input_df = pd.DataFrame()
        self.input_df_ok = False
        self.input_df_errors = []
        self.parents_in_df_ok = False
        self.parents_email_errors = []
        self.parent_email_map = {}
        self.display_button_timestamp = time.time()
        self.send_emails_button_timestamp = time.time()

    def create_dashboard(self):

        self.app.layout = html.Div(children=[
            dbc.Row([
                dbc.Col([
    html.H1(children='Velkommen til Mathias Gran sitt epost-program'),
                    html.Br(),
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    'Dra excel-filen hit eller ',
                    html.A('trykk for å velge fil')
                ]),
                style={
                    'width': '20%',
                    'height': 'auto',
                    'lineHeight': '20px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin': '10px'
                }),
            html.Div(id='output-data-upload', children=["Venter på fil"]),
                    html.Div(id='output-data-upload2', children=[""])
                    ], width=8), dbc.Col([
                    html.Img(src=self.app.get_asset_url('Sheep.JPG'), width="60%", height="auto")
                    ], width=4)
            ]),
            html.Hr(),



            html.Div(children=[
                dbc.Row([
                    dbc.Col([
                        dcc.Dropdown(id='recievers_dropdown', clearable=False),
                        html.Br(),
                        html.P("Kodene for å mappe til de ulike inputene:"),
                        html.Div(id="Mappings", style={'border': '2px solid black'}),
                        html.Br(),
                        html.P("Skriv inn teksten til emnefeltet:"),
                        dcc.Input(id='Header'),
                        html.Br(),
                        html.Div([
                            html.Br(),
                            html.Br(),
                            dbc.Button("Se eksempel på epost", color="dark", className="me-1", id="Display_email_button"),
                            html.Div(id="button1_out"),
                            html.Br(),
                            html.Br(),
                            dbc.Button("Generer og send epost til alle", color="warning", className="me-1", id="Send_emails_button"),
                            html.Div(id="button2_out"),
                        ], id="Execute_area", style={'display': 'none'})
                        ], width=4),

                    dbc.Col([
                        html.P("Skriv epost med mappingen her:"),
                        dcc.Textarea(id="Editable textarea", readOnly=False, style={'width': '100%', 'height': 300})
                    ], width=4),
                    dbc.Col([
                        html.P("Se hvordan eposten vil se ut her:"),
                        dcc.Textarea(id="Non-editable textarea", readOnly=True, style={'width': '100%', 'height': 300})
                    ], width=4)
            #html.Div(id="test_output")
        ])
        ], id="EmailMakerDiv", style={"display": 'none'})
        ])



        @self.app.callback([Output('output-data-upload', 'children'),
                            Output('output-data-upload2', 'children')],
                      Input('upload-data', 'contents'),
                      State('upload-data', 'filename'),
                      State('upload-data', 'last_modified'))
        def update_output(excel_file, name, date):
            if excel_file is not None:
                df = self.read_excel_content(excel_file)
                self.check_excel_file(df)

                self.input_df = df

                if self.input_df_ok:
                    return_1 = "Fil lastet opp"

                    if self.parents_in_df_ok:
                        return_2 = ""
                    else:
                        return_2 = [html.P("Funksjonalitet for å sette foreldre på kopi er avslått grunnet følgende feil:")]
                        for error in self.parents_email_errors:
                            return_2.append(html.P(error))

                else:
                    return_1 = [html.P("Filen ble ikke lastet opp korrekt. Følgende feil eksisterer:")]
                    for error in self.input_df_errors:
                        return_1.append(html.P(error))

                return return_1, return_2

            else:
                return "", ""


        @self.app.callback([Output('recievers_dropdown', 'options'),
                            Output('recievers_dropdown', 'value')],
                           [Input("output-data-upload", 'children')],
                           State("output-data-upload", 'children'))
        def update_dropdown(upload_text, state):
            if upload_text == "Fil lastet opp":
                options = [{'label': i, 'value': i} for i in self.input_df['Epost']]
                return options, options[0]["label"]
            else:
                return [], ""

        @self.app.callback(Output('Mappings', 'children'),
                           Input('recievers_dropdown', 'value'))
        def update_mapping(email):
            mapping_list = []
            for col in self.input_df.columns:
                if col not in ["Epost", "Født", "Epost_foresatte"]:
                    mapping_list.append(html.P(
                        f" {col} = <{col}>, som for denne mailadressen er {self.input_df[col].loc[self.input_df['Epost'] == email].iloc[0]}. <{col}!> for små bokstaver"))

            return mapping_list

        @self.app.callback(Output("Non-editable textarea", 'value'),
                           [Input("Editable textarea", 'value'),
                           Input('recievers_dropdown', 'value')])
        def show_email_text(text, email):
            return self.write_email_text(email, text)

        @self.app.callback(Output("button1_out", 'children'),
                            [Input('Editable textarea', 'value'),
                            Input('Header', 'value'),
                            Input('recievers_dropdown', 'value'),
                            Input('Display_email_button', 'n_clicks_timestamp')])
        def see_example_email(text, header, email, button_timestamp):
            if button_timestamp is not None:
                if button_timestamp > self.display_button_timestamp:
                    self.display_button_timestamp = button_timestamp
                    text = self.write_email_text(email, text)
                    if self.generate_email(email, header, text, send=False):
                        return html.P("Eksempel på epost generert")


        @self.app.callback(Output("button2_out", 'children'),
                           [Input('Editable textarea', 'value'),
                            Input('Header', 'value'),
                            Input('Send_emails_button', 'n_clicks_timestamp')])
        def send_emails(text, header, button_timestamp):
            if button_timestamp is not None:
                if button_timestamp > self.send_emails_button_timestamp:
                    self.send_emails_button_timestamp = button_timestamp
                    for email in list(self.input_df["Epost"]):
                        user_text = self.write_email_text(email, text)
                        self.generate_email(email, header, user_text, send=True)

                    return html.P("Eposter sendt!")


        @self.app.callback(Output('EmailMakerDiv', 'style'),
                           [Input("output-data-upload", 'children')])
        def show_input_area(input_text):
            if input_text == "Fil lastet opp":
                return {"display": 'block'}
            else:
                return {"display": 'none'}

        @self.app.callback(Output('Execute_area', 'style'),
                           [Input('Non-editable textarea', 'value'),
                            Input('Header', 'value')])
        def show_execute_area(text, header):
            if text is None:
                text = ""
            if header is None:
                header = ""

            if len(text) > 5 and len(header) > 3:
                return {"display": 'block'}
            else:
                return {"display": 'none'}



    def generate_email(self, email, header, text, send=False):
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = email
            mail.Subject = header
            mail.HTMLBody = f'<h3>{text}</h3>'
            mail.Body = text

            if self.parents_in_df_ok:
                if email in self.parent_email_map:
                    mail.CC = self.parent_email_map[email]

            if send:
                mail.send()
            else:
                mail.display()

            return True
        except Exception as e:
            print(e)
            return False


    def write_email_text(self, email, text):
        if text is not None:
            for col in self.input_df.columns:
                if col not in ["Epost", "Født", "Epost_foresatte"]:
                    text = text.replace(f"<{col}>", self.input_df[col].loc[self.input_df['Epost'] == email].iloc[0])
                    text = text.replace(f"<{col}!>", (self.input_df[col].loc[self.input_df['Epost'] == email].iloc[0]).lower())

            return text


    def check_excel_file(self, df):

        self.parents_email_errors = []
        self.input_df_errors = []

        df_col = list(df.columns)
        if not "Epost" in df_col:
            self.input_df_errors.append("Epost er ikke en kolonne i inputfil")

        elif not len(list(df["Epost"])) == len(set(df["Epost"])):
            duplicate_dict = Counter(list(df["Epost"]))
            for email in duplicate_dict:
                if duplicate_dict[email] > 1:
                    self.input_df_errors.append(f"{email} fremkommer flere ganger i inputfil")

        else:
            counter = 0
            for email in list(df["Epost"]):
                counter += 1
                if email is None:
                    self.input_df_errors.append(f"Epost nummer {counter} fra toppen er blank")

                elif "@" not in email or "." not in email:
                    self.input_df_errors.append(f"{email} er ikke en gyldig epost")

        if len(self.input_df_errors) > 0:
            self.input_df_ok = False
        else:
            self.input_df_ok = True

        if not("Født" in df_col and "Epost_foresatte" in df_col):
            self.parents_email_errors.append("Inputfil mangler Født eller Epost_foresatte som kolonne")
        else:
            today = datetime.datetime.now()
            for email in list(df["Epost"]):
                if df["Født"].loc[df["Epost"]==email].iloc[0] is None:
                    self.parents_email_errors.append(f"{email} mangler fødselsdato")

                else:
                    try:
                        birth_string = df["Født"].loc[df["Epost"]==email].iloc[0]
                        birth_date = datetime.datetime.strptime(birth_string, "%d.%m.%Y")

                        try:
                            adult_age = birth_date.replace(year=birth_date.year + 18)
                        except ValueError:
                            # 👇️ preserve calendar day (if Feb 29th doesn't exist, set to 28th)
                            adult_age = birth_date.replace(year=birth_date.year + 18, day=28)

                        #print(f"{today.strftime('%d.%m.%Y')} & {adult_age.strftime('%d.%m.%Y')}")
                        is_adult = adult_age <= today


                    except:
                        self.parents_email_errors.append(f"{email} sin fødselsdato er på feil format. Korrekt format er (dd.mm.åååå)")
                        is_adult = True

                    if not is_adult:
                        if type(df["Epost_foresatte"].loc[df["Epost"]==email].iloc[0])==float:
                            self.parents_email_errors.append(f"{email} er under 18 år, men har ikke en registrert foresatt-epost")
                        elif "@" not in str(df["Epost_foresatte"].loc[df["Epost"]==email].iloc[0]) or "." not in str(df["Epost_foresatte"].loc[df["Epost"]==email].iloc[0]):
                            self.parents_email_errors.append(f"{email} er under 18 år, men registrert foresatt-epost er på feil format")
                        else:
                            self.parent_email_map[email] = df["Epost_foresatte"].loc[df["Epost"]==email].iloc[0]

        if len(self.parents_email_errors)>0:
            self.parents_in_df_ok = False
        else:
            self.parents_in_df_ok = True




    def read_excel_content(self, contents):
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        df = pd.read_excel(io.BytesIO(decoded), sheet_name="Input")
        return df

    def run_app(self):
        url = "http://127.0.0.1:8051/"
        webbrowser.open_new_tab(url)
        self.app.run_server(debug=False, port=8051)

if True:
    app_object = DashboardMaker()
    app_object.create_dashboard()
    app_object.run_app()