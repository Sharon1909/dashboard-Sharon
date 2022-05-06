import pandas as pd  
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="GIP Dashboard", 
    page_icon=":bar_chart:", 
    layout="wide",
    initial_sidebar_state="expanded"
    )

# ---- MAINPAGE ----
st.title(":bar_chart: Jaarrekening dashboard Baltimore Aircoil Company")
st.markdown("##")

from PIL import Image
image = Image.open("BAC.jpg")
st.image(image)

st.subheader("Kies een van de ratio's in de zijbalk.")

# ---- SIDEBAR ----
st.sidebar.header("Gelieve hier te filteren:")
SelectedCategory = st.sidebar.multiselect(
    "Selecteer de ratio die je wenst:",
    ("Samenstelling activa & passiva","Liquiditeit","Solvabiliteit","Rentabiliteit","Omlooptijd voorraden","Soorten voorraden")
)
boekjaar = st.sidebar.radio("Selecteer boekjaar:",
("Boekjaar 1","Boekjaar 2","Boekjaar 3"),index=0)

# ---- READ EXCEL ACTIVA ----
@st.cache
def get_activa_from_excel():
        df = pd.read_excel(
            io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
            engine="openpyxl",
            sheet_name="verticale analyse balans",
            usecols="A:E",
            nrows=100,
            header=2
        )

        # filter row on column value
        activa = ["VASTE ACTIVA","VLOTTENDE ACTIVA"]
        df = df[df['ACTIVA'].isin(activa)]

        return df

df_activa = get_activa_from_excel()

# ---- READ EXCEL PASIVA ----
@st.cache
def get_passiva_from_excel():
        df = pd.read_excel(
            io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
            engine="openpyxl",
            sheet_name="verticale analyse balans",
            usecols="A:E",
            nrows=100,
            header=49
        )

        # filter row on column value
        passiva = ["EIGEN VERMOGEN","VOORZIENINGEN EN UITGESTELDE BELASTINGEN","SCHULDEN"]
        df = df[df['PASSIVA'].isin(passiva)]

        return df

df_passiva = get_passiva_from_excel()

#---- READ EXCEL REV ----
def get_rev_from_excel():
        df = pd.read_excel(
            io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
            engine="openpyxl",
            sheet_name="REV",
            usecols="A:D",
            nrows=10,
            header=1
        )
        # change column names
        df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
        # filter row on column value
        rev = ["REV"]
        df = df[df["Type"].isin(rev)]

        df = df.T #Transponeren
        df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                        "Boekjaar 3":"3"})
        df = df.iloc[1: , :] # Drop first row 
        df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                        "Boekjaar 3"],True)
        df.columns = ["Boekjaar","REV"] # change column names
        
        
        return df

df_rev = get_rev_from_excel()

#---- READ EXCEL LIQ ---
def get_liq_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
        engine="openpyxl",
        sheet_name="Liquiditeit",
        usecols="A:D",
        nrows=32,
        header=1
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    liq = ["Liquiditeit in ruime zin","Liquiditeit in enge zin"]
    df = df[df["Type"].isin(liq)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                        "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                        "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Liquiditeit in ruime zin","Liquiditeit in enge zin"] # change column names
        
        
    return df

df_liq = get_liq_from_excel()

#---- READ EXCEL SOLV ----
def get_solv_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
        engine="openpyxl",
        sheet_name="Solvabiliteit",
        usecols="A:D",
        nrows=6,
        header=0
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    solv = ["Solvabiliteit"]
    df = df[df["Type"].isin(solv)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                        "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                        "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Solvabiliteit"] # change column names
        
        
    return df

df_solv = get_solv_from_excel()

#---- READ EXCEL OMLOOPTIJD VOORRAAD ----
def get_omvoorraad_from_excel():
    df = pd.read_excel(
        io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
        engine="openpyxl",
        sheet_name="Voorraad",
        usecols="A:D",
        nrows=6,
        header=0
    )
    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    voorraad = ["Omlooptijd"]
    df = df[df["Type"].isin(voorraad)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                        "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                        "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Voorraad"] # change column names
        
        
    return df

df_omvoorraad = get_omvoorraad_from_excel()

#---- READ EXCEL VOORRADEN ----
@st.cache
def get_voorraden_from_excel():
    df = pd.read_excel(
    io="data/GIP_analyse van de jaarrekening_SharonSerneels.xlsx",
    engine="openpyxl",
    sheet_name="verticale analyse balans",
    usecols="A:E",
    nrows=100,
    header=2
    )

    # change column name
    df.columns = ["Type", "Codes", "Boekjaar 1", "Boekjaar 2", "Boekjaar 3"]

    # filter row on column value
    voorraden = ["Grond- en hulpstoffen","Goederen in bewerking", "Gereed product", "Handelsgoederen"]
    df = df[df['Type'].isin(voorraden)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2","Boekjaar 3":"3"})
    df = df.iloc[2: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2","Boekjaar 3"],True)
    df.columns = ["Boekjaar","Grond- en hulpstoffen","Goederen in bewerking", "Gereed product", "Handelsgoederen"] # change column names
    df = df.astype({"Grond- en hulpstoffen":"float64","Goederen in bewerking":"float64", "Gereed product":"float64", "Handelsgoederen":"float64"})
        
        
    return df

df_voorraden = get_voorraden_from_excel()

for chart in SelectedCategory:
    if chart == "Samenstelling activa & passiva":
        st.subheader("Pas in de zijbalk aan welk boekjaar je wilt bekijken.")

        #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='activa'):
            st.subheader('Activa')
            st.write(df_activa)

        fig_activa = px.pie(df_activa, 
                values=boekjaar, 
                names='ACTIVA',
                title=f'Samenstelling activa {boekjaar}',            
                color_discrete_sequence = ['#0DA2FF', '#52B2BF'])

        fig_activa.update_traces(textfont_size=15, pull=[0, 0.2], marker=dict(line=dict(color='#000000', width=2)))
        fig_activa.update_layout(legend = dict(font = dict(size = 15)), title = dict(font = dict(size = 25)))
        st.plotly_chart(fig_activa, use_container_width=True)

        #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='passiva'):
            st.subheader('Passiva')
            st.write(df_passiva)

        fig_passiva = px.pie(df_passiva,
            values=boekjaar, 
            names='PASSIVA',
            title= f'Samenstelling passiva {boekjaar}',
            color='PASSIVA',
            color_discrete_sequence = ['#0DA2FF', '#86DC3D', '#52B2BF'])

        fig_passiva.update_traces(textfont_size=15, pull=[0, 0.2], marker=dict(line=dict(color='#000000', width=2)))
        fig_passiva.update_layout(legend = dict(font = dict(size = 15)), title = dict(font = dict(size = 25)))
        st.plotly_chart(fig_passiva, use_container_width=True)

    elif chart =="Rentabiliteit":
        #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='rentabiliteit'):
            st.subheader('Rentabiliteit')
            st.write(df_rev)

        fig_rev = px.line(df_rev, x="Boekjaar", y="REV", markers=True)
        fig_rev.update_layout({
        'plot_bgcolor': 'rgba(0, 0, 0, 0)',
        'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

        fig_rev.update_traces(line=dict(width=3),line_color = '#0DA2FF')
        fig_rev.update_layout(title_text='Rentabiliteit van het EV', title_x=0.5)
        st.plotly_chart(fig_rev, use_container_width=True)

    elif chart == "Liquiditeit":
    #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='liquiditeit'):
            st.subheader('Liquiditeit')
            st.write(df_liq)

        fig_liq = px.line(df_liq, x="Boekjaar",
        y=["Liquiditeit in ruime zin","Liquiditeit in enge zin"],
        markers=True,
        color_discrete_sequence = ['#0DA2FF', '#86DC3D'])
        
        fig_liq.update_layout({
        'plot_bgcolor': 'rgba(0, 0, 0, 0)',
        'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

        fig_liq.update_traces(line=dict(width=3))
        fig_liq.update_layout(title_text='Liquiditeit', title_x=0.5, xaxis_title="Boekjaar", yaxis_title="Liquiditeit")
        st.plotly_chart(fig_liq, use_container_width=True)

    elif chart == "Solvabiliteit":
        #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='solvabiliteit'):
            st.subheader('Solvabiliteit')
            st.write(df_solv)

        fig_solv = px.line(df_solv, x="Boekjaar", y="Solvabiliteit", markers=True)
        fig_solv.update_layout({
        'plot_bgcolor': 'rgba(0, 0, 0, 0)',
        'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

        fig_solv.update_traces(line=dict(width=3), line_color="#0DA2FF")
        fig_solv.update_layout(title_text='Solvabiliteit', title_x=0.5)
        st.plotly_chart(fig_solv, use_container_width=True)

    elif chart == "Omlooptijd voorraden":
        #Use a button to toggle data
        if st.checkbox('Zie cijfers:', key='om_voorraad'):
            st.subheader('Omlooptijd Voorraad')
            st.write(df_omvoorraad)

        fig_omvoorraad = px.bar(df_omvoorraad, x="Boekjaar", y="Voorraad")
        
        fig_omvoorraad.update_layout({
        'plot_bgcolor': 'rgba(0, 0, 0, 0)',
        'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

        fig_omvoorraad.update_traces(width=0.40, marker_color='#0DA2FF')
        fig_omvoorraad.update_layout(title_text='Omlooptijden van de voorraad', title_x=0.5, xaxis_title="Boekjaar", yaxis_title="Omlooptijd")
        st.plotly_chart(fig_omvoorraad, use_container_width=True)

    elif chart == "Soorten voorraden":
        if st.checkbox('Zie cijfers:', key='voorraden'):
            st.subheader('Soorten voorraad')
            st.write(df_voorraden)

        fig_voorraden = px.bar(df_voorraden, x="Boekjaar", 
        y=["Grond- en hulpstoffen",'Goederen in bewerking','Gereed product','Handelsgoederen'],
        color_discrete_sequence=["rgb(13, 162, 255)", "rgb(134, 220, 61)", "rgb(82, 178, 191)", "rgb(136, 199, 220)"])

        fig_voorraden.update_layout({
        'plot_bgcolor': 'rgba(0, 0, 0, 0)',
        'paper_bgcolor': 'rgba(0, 0, 0, 0)',})

        fig_voorraden.update_traces(width=0.6)
        fig_voorraden.update_layout(title_text='Soorten voorraden', title_x=0.5, yaxis_range=[0,0.13])
        st.plotly_chart(fig_voorraden, use_container_width=True)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)