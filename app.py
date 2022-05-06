from email import header
from tkinter import Variable
import pandas as pd  
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="GIP Dashboard", 
    page_icon=":bar_chart:", 
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
         'Get Help': 'https://www.extremelycoolapp.com/help',
         'Report a bug': "https://www.extremelycoolapp.com/bug",
         'About': "# This is a header. This is an *extremely* cool app!"}
    )

# ---- READ EXCEL Liquiditeit ruime zin ----
@st.cache
def get_liquiditeit_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipgeg2liq",
        usecols="A:D",
        nrows=15, 
        header=1
        
    )

    # filter row on column value
    df.columns = ['Vlottende activa','Boekjaar 1','Boekjaar 2','Boekjaar 3']
    vlot= ['Voorraden en bestellingen in uitvoering','Vorderingen op ten hoogste één jaar','Geldbeleggingen','Liquide middelen','Overlopende rekeningen','Totaal']
    df= df[df['Vlottende activa'].isin(vlot)]
    return df

df_vlot = get_liquiditeit_from_excel()
# ---- READ EXCEL Liquiditeit enge zin ----
@st.cache
def get_liquiditeit_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipgeg3liq",
        usecols="A:D",
        nrows=15, 
        header=1
        
    )

    # filter row on column value
    df.columns = ['Vlottende activa','Boekjaar 1','Boekjaar 2','Boekjaar 3']
    vlota= ['Vorderingen op ten hoogste één jaar','Geldbeleggingen','Liquide middelen','Overlopende rekeningen','Totaal']
    df= df[df['Vlottende activa'].isin(vlota)]
    return df

df_vlota = get_liquiditeit_from_excel()
# ---- READ EXCEL Liquiditeit ruime zin schuld ----
@st.cache
def get_liquiditeitschuld_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipgegliq",
        usecols="A:D",
        nrows=7,
        header=1
    )

    # filter row on column value
    liquiruimschulds = ["Schulden op ten hoogste één jaar","Overlopende rekeningen","Totaal"]
    df = df[df['Vreemd vermogen op KT'].isin(liquiruimschulds)]
    return df

df_liquiruimschuld = get_liquiditeitschuld_from_excel()

# ---- READ EXCEL Liquiditeit enge zin schuld ----
@st.cache
def get_liquiditeitengeschuld_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipgeg4liq",
        usecols="A:D",
        nrows=7,
        header=1
    )

    # filter row on column value
    liquiruimschuldseng = ["Schulden op ten hoogste één jaar","Overlopende rekeningen","Totaal"]
    df = df[df['Vreemd vermogen op KT'].isin(liquiruimschuldseng)]
    return df

df_liquiruimschuldeng = get_liquiditeitengeschuld_from_excel()

# ---- READ EXCEL Liquiditeit resultaat----
@st.cache
def get_liquiditeitbenodig_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipgeg7liq",
        usecols="A:D",
        nrows=20,
        header=2
    )

    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    liquibok = ["Liquiditeit in ruime zin","Liquiditeit in enge zin"]
    df = df[df["Type"].isin(liquibok)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Liquiditeit in ruime zin","Liquiditeit in enge zin"] # change column names
    
    
    return df
df_liquinodigheid= get_liquiditeitbenodig_from_excel()

# grafiek liquiditeit
fig_liquiditeit = px.line(df_liquinodigheid,x='Boekjaar',  y=["Liquiditeit in ruime zin","Liquiditeit in enge zin"], markers=True ,labels={
                     'value': "liquiditeit",'variable':"Liquiditeiten"},title=('Liquiditeit grafiek') )        
fig_liquiditeit.update_layout({
'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',},title = dict(font = dict(size = 30)))
fig_liquiditeit.update_traces(line=dict(width=3))


# ---- READ EXCEL Solvabiliteit resultaat----
@st.cache
def get_solvanodig_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipsolva1",
        usecols="A:D",
        nrows=10,
        header=2
    )

    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    solvabok = ["Solvabiliteit"]
    df = df[df["Type"].isin(solvabok)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Solvabiliteit"] # change column names
    
    
    return df
df_solvanodigheid= get_solvanodig_from_excel()

# ---- READ EXCEL Solvabiliteit gegevens uit bestaan ----
@st.cache
def get_Solvage_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Gipsolva1",
        usecols="A:D",
        nrows=10,
        header=2
    )

    # filter row on column value
    Solva = ["Eigen vermogen","Totaal vermogen"]
    df = df[df['Boekjaar'].isin(Solva)]
    return df

df_Solvageg = get_Solvage_from_excel()

# grafiek Solvabiliteit
fig_solvabilie = px.bar(df_solvanodigheid, x='Boekjaar',y="Solvabiliteit",
                title=('Solvabiliteit grafiek'))
fig_solvabilie.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',},title = dict(font = dict(size = 30)))
# ---- READ EXCEL Rendabiliteit benodigheid----
@st.cache
def get_Rendanodig_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Revgeg2",
        usecols="A:D",
        nrows=10,
        header=3
    )

    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    Rendabok = ["REV"]
    df = df[df["Type"].isin(Rendabok)]

    df = df.T #Transponeren
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","REV"] # change column names
    
    
    return df
df_rendanodigheid= get_Rendanodig_from_excel()
# ---- READ EXCEL Rendabiliteit gegevens uit bestaan ----
@st.cache
def get_Rendabil_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="Revgeg1",
        usecols="A:D",
        nrows=16,
        header=2
    )

    # filter row on column value
    Rev = ["Winst van het boekjaar na belastingen","Eigen Vermogen"]
    df = df[df['Boekjaar'].isin(Rev)]
    return df

df_Revgeg = get_Rendabil_from_excel ()    

# rendabiliet grafiek
fig_Renda = px.area(df_rendanodigheid, x="Boekjaar", y="REV",labels={
                     'value': "Rendabiliteit",'variable':"Rendabiliteiten"},title=('Rendabiliteit grafiek'))
fig_Renda.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',},title = dict(font = dict(size = 30)))
# ---- READ EXCEL KlantLevkrediet gegevens uit bestaan ----
@st.cache
def get_Klantlevkrediet_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="KlantLevKrediet",
        usecols="A:D",
        nrows=5,
        header=1
    )

    # filter row on column value
    Rev = ["handelsvorderingen (code 40)","omzet (code 70) inclusief btw code 9146"]
    df = df[df['Boekjaar'].isin(Rev)]
    return df

df_Klantkrediet = get_Klantlevkrediet_from_excel ()  
# ---- READ EXCEL KlantLevkrediet gegevens uit bestaan ----
@st.cache
def get_Levkrediet_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="KlantLevKrediet",
        usecols="A:D",
        nrows=10,
        header=1
    )

    # filter row on column value
    Rev = ["handelsschulden 44","aankopen (code 600/8+code61) incl btw code 9145"]
    df = df[df['Boekjaar'].isin(Rev)]
    return df

df_LevKrediet = get_Levkrediet_from_excel ()   
# ---- READ EXCEL KlantLevkrediet benodigheid----
@st.cache
def get_KLnodig_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="KlantLevKrediet",
        usecols="A:D",
        nrows=10,
        header=1
    )    

    # change column names
    df.columns = ["Type","Boekjaar 1","Boekjaar 2","Boekjaar 3"]
    # filter row on column value
    Klev = ["Klantenkrediet","Leverancierskrediet"]
    df = df[df['Type'].isin(Klev)]
    
    df = df.T #Transponeren
    
    df = df.rename(index={"Boekjaar 1":"1","Boekjaar 2":"2",
                    "Boekjaar 3":"3"})
    df = df.iloc[1: , :] # Drop first row 
    df.insert(0,"Boekjaar",["Boekjaar 1","Boekjaar 2",
                    "Boekjaar 3"],True)
    df.columns = ["Boekjaar","Klantenkrediet","Leverancierskrediet"] # change column names
    df = df.astype({'Klantenkrediet': 'float64','Leverancierskrediet': 'float64'})
    return df
df_KlantLevnodigheid= get_KLnodig_from_excel()
#grafiek Leveranier en klant
fig_klantLev = px.bar(df_KlantLevnodigheid, x=["Klantenkrediet","Leverancierskrediet"],
    y="Boekjaar",
    barmode="group",
    orientation="h",labels={ 'value': "Aantal dagen",'variable':"Soorten krediet"},title=('KlantLevkrediet grafiek'))
fig_klantLev.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)',
'paper_bgcolor': 'rgba(0, 0, 0, 0)',},title = dict(font = dict(size = 30)))
# ---- READ EXCEL ACTIVA ----
@st.cache
def get_activa_from_excel():
    df = pd.read_excel(
        io="data/Varo oplossing.xlsx",
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
        io="data/Varo oplossing.xlsx",
        engine="openpyxl",
        sheet_name="verticale analyse balans",
        usecols="A:E",
        nrows=100,
        header=50
    )

    # filter row on column value
    passiva = ["EIGEN VERMOGEN","VOORZIENINGEN EN UITGESTELDE BELASTINGEN","SCHULDEN"]
    df = df[df['PASSIVA'].isin(passiva)]

    return df

df_passiva = get_passiva_from_excel()

# ---- MAINPAGE ----
st.title(":bar_chart: Jaarrekening Dashboard")
st.markdown("##")
# ---- SIDEBAR ----
st.sidebar.header("Gelieve hier te filteren:")
boekjaar = st.sidebar.radio(
    "Selecteer boekjaar:",
    ("Boekjaar 1","Boekjaar 2","Boekjaar 3"),
    index=0
)

# ---- SIDEBAR grafiek ----
grafiek = st.sidebar.radio(
    "Selecteer een ratio:",
    ("Liquiditeit","Rendabiliteit","Solvabiliteit","Klantlevkrediet"),
    index=0
)

if grafiek == 'Liquiditeit':
    st.header('Liquiditeit')
    st.plotly_chart(fig_liquiditeit, use_container_width=True)
    agree = st.checkbox('Showe')
    if agree:
        col1, col2 = st.columns([1,1]) 
        with col1:
            st.write(df_vlot)
        with col2:
            st.write(df_liquiruimschuld)
elif grafiek == 'Rendabiliteit':
    st.header('Rendabiliteit')
    st.plotly_chart(fig_Renda, use_container_width=True)
    agree = st.checkbox('Showe')
    if agree:     
        st.write(df_Revgeg)
elif grafiek == 'Solvabiliteit':
    st.header('Solvabiliteit')
    st.plotly_chart(fig_solvabilie, use_container_width=True)
    agree = st.checkbox('Showe')
    if agree: 
        st.write(df_solvanodigheid)
elif grafiek == 'Klantlevkrediet':
    st.header('KlantLevKrediet')
    st.plotly_chart(fig_klantLev, use_container_width=True)
    agree = st.checkbox('Showe')
    if agree:
        col1, col2 = st.columns([1,1]) 
        with col1:
            st.write(df_Klantkrediet)
        with col2:
            st.write(df_LevKrediet)

# Samenstelling activa boekjaar [TAART DIAGRAM]
fig_activa = px.pie(df_activa, 
            values=boekjaar, 
            names='ACTIVA',
            title=f'Samenstelling activa {boekjaar}'            
            )
fig_activa.update_traces(textfont_size=20, pull=[0, 0.2], marker=dict(line=dict(color='#000000', width=2)))
fig_activa.update_layout(legend = dict(font = dict(size = 20)), title = dict(font = dict(size = 30)))

# Samenstelling pasiva boekjaar [TAART DIAGRAM]
fig_passiva = px.pie(df_passiva, 
            values=boekjaar, 
            names='PASSIVA',
            title= f'Samenstelling passiva {boekjaar}'            
            )
fig_passiva.update_traces(textfont_size=20, pull=[0, 0.2], marker=dict(line=dict(color='#000000', width=2)))
fig_passiva.update_layout(legend = dict(font = dict(size = 20)), title = dict(font = dict(size = 30)))


##test
#
#
#left_column, right_column = st.columns(2)
col1, col2 = st.columns([1,1])
with col1:
    st.write(df_activa)
    st.plotly_chart(fig_activa, use_container_width=True)
with col2:
    st.write(df_passiva)
    st.plotly_chart(fig_passiva, use_container_width=True)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)