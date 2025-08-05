import streamlit as st
import pandas as pd
import regex as re

st.set_page_config(page_title="OMS naar OHJP Conversie Tool", layout="wide")

@st.cache_data
def load_excel(file, header=0):
    try:
        df = pd.read_excel(file, engine='openpyxl', header=header)
        return df
    except Exception as e:
        st.error(f"Fout bij het laden van het Excel-bestand: {e}")
        return None
    
def normalize_frequency(row):
    """Normalize the frequency of the tasks in the DataFrame.
    Frequencies "WEEKS" and "YEARS" are converted to "MONTHS".
    Column "Frequentie aantal" is changed using multiplication or division.
    
    :param row: A row of the DataFrame.
    :return: The modified row with normalized frequency.
    """

    if row["Eenheid"] == "WEEKS":
        row["Interval"] = row["Interval"] * 0.25
        row["Eenheid"] = "MONTHS"
    elif row["Eenheid"] == "YEARS":
        row["Interval"] = row["Interval"] * 12
        row["Eenheid"] = "MONTHS"
    elif row["Eenheid"] == "DAYS":
        if row["Interval"] == 365:
            row["Interval"] = 12
        else:
            row["Interval"] = (round((row["Interval"] / 30),2))
            row["Eenheid"] = "MONTHS"
    return row

def extract_uitvoerende(row):
    """Extract the executing party from the 'Omschrijving' and 'Taakplan omschrijving' column.

    Possible key words in string:

    - LEV : Leverancier
    - OA : Onderaannemer
    
    :param row: A row of the DataFrame.
    :return: The modified row with the executing party extracted.
    """
    if len(str(row['Omschrijving'])) == 0 or len(str(row['Taakplan omschrijving'])) == 0: # If one of the columns is empty, set 'Uitvoerende' to empty string
        row['Uitvoerende'] = ''
        return row
    if re.search(r'\bLEV\b', row['Omschrijving'], flags=re.IGNORECASE) or re.search(r'\bLEV\b', row['Taakplan omschrijving'], flags=re.IGNORECASE):
        row['Uitvoerende'] = 'Leverancier'
    elif re.search(r'\bOA\b', row['Omschrijving'], flags=re.IGNORECASE) or re.search(r'\bOA\b', row['Taakplan omschrijving'], flags=re.IGNORECASE):
        row['Uitvoerende'] = 'Onderaannemer'

    return row

def filter_columns(df):
    """Filter the DataFrame to only include the necessary columns for OHJP.

    :param df: The DataFrame to filter.
    :return: The filtered DataFrame.
    """
    column_list = ['Type', 'Omschrijving', 'Taakplan omschrijving',
       'Locatie omschrijving', 'Startdatum', 'Einddatum', 'Startdatum wk',
       'Einddatum wk', 'PMnum', 'Interval', 'Eenheid', 'Nummer', 'Taakplannr.']
    #If one of the columns is not present in the DataFrame, return an error message
    for column in column_list:
        if column not in df.columns:
            st.error(f"Kolom '{column}' ontbreekt in het bestand. Zorg ervoor dat je een geldig OMS-exportbestand hebt geüpload.")
            return None
    df = df[column_list]
    df = df.rename(columns={"Locatie omschrijving": "Object"})
    return df

def create_heatmap_df(df, start_week=1):
    """
    Create a heatmap dataframe with unique values under "Omschrijving" as index and week numbers as columns.
    The values in the dataframe are >= 1 if the planning block is present, otherwise it will be 0.
    All week numbers are present in the columns.
    """

    # Create week nr columns 1 to 52
    week_numbers = list(range(1, 53))
    for week in week_numbers:
        df[str(week)] = 0


    # Create a pivot table with "Omschrijving" as index, "Week" as columns, and count of occurrences as values
    heatmap_df = df.pivot_table(index=['PMnum', 'Interval'], columns='Week', aggfunc='size', fill_value=0)
    heatmap_df = heatmap_df.reset_index()

    #Also add week numbers that had no values
    for week in week_numbers:
        if week not in heatmap_df.columns:
            heatmap_df[week] = 0
    
    #Find metadata (Frequentie aantal, Frequentie, Uitvoerende) and add it to the heatmap_df. Using the first occurrence of the metadata in the original dataframe. Add the metadata to the heatmap_df as new columns.
    metadata = df[['Complex', 'Omschrijving', 'Taakplan omschrijving', 'Object', 'Interval', 'Eenheid', 'PMnum', 'Uitvoerende', 'Nummer', 'Taakplannr.', 'Route']].drop_duplicates(subset=['PMnum', 'Interval'])
    

    heatmap_df = heatmap_df.merge(metadata, on=['PMnum', 'Interval'], how='left')

    #sort week_numbers starting with start_week to 52, then 1 to start_week-1
    week_numbers = [week for week in week_numbers[start_week-1:] + week_numbers[:start_week-1]]
    # Sort columns: metadata first, then week numbers
    cols = ["Complex", "Object", 'Omschrijving', 'Interval', 'Uitvoerende', 'Taakplan omschrijving', 'Nummer', 'Taakplannr.', 'Route'] + week_numbers
    heatmap_df = heatmap_df[cols]

    return heatmap_df

def adapt_to_version(heatmap_df, version=1):
    """Adapt the code to the latest version of the OMS to OHJP conversion tool."""
    # This function can be used to adapt the code to the latest version of the OMS to OHJP conversion tool.
    if version == 1:
        # Remove columns 'Nummer' and 'Taakplannr.' if they exist
        if 'Nummer' in heatmap_df.columns:
            heatmap_df = heatmap_df.drop(columns=['Nummer'])
        if 'Taakplannr.' in heatmap_df.columns:
            heatmap_df = heatmap_df.drop(columns=['Taakplannr.'])
        if 'Route' in heatmap_df.columns:
            heatmap_df = heatmap_df.drop(columns=['Route'])
    return heatmap_df


st.title("🗓️ OMS naar OHJP Conversie Tool")

st.write("Deze tool helpt jou met het omzetten van een OMS-export uit Maximo naar een OHJP-worksheet" \
" voor het opzetten van een onderhoudsjaarplan.")

with st.sidebar:
    st.markdown("**1. Upload een OMS-exportbestand (Excel-format) om te beginnen:** 👇")
    uploaded_file = st.file_uploader("Kies een bestand", type=["xlsx"])
    header_row = st.number_input("Kies de rij waar de kolomnamen staan in je Excel (standaard is 0)", min_value=0, value=0)
    if uploaded_file is not None:
        df = load_excel(uploaded_file, header=header_row)
        if df is not None:
            st.session_state['df'] = df
    # Vraag gebruiker om een lijst te uploaden om objecten in de upload te matchen met een complex

    st.markdown("**2. Upload een lijst met objecten om te matchen met een complex:** 👇")
    file_object_complex = st.file_uploader("Kies een bestand met objecten. De kolomnamen moeten staan in de bovenste rij van het Excel-bestand.", type=["xlsx"])
    if file_object_complex is not None:
        df_object_complex = load_excel(file_object_complex, header=0)
        # Ask user to select the column with the object names
        object_name_column = st.selectbox("Kies de kolom met objectnamen:", df_object_complex.columns, index=None)
        # Ask user to select the column with the complex names
        complex_name_column = st.selectbox("Kies de kolom met complexnamen:", df_object_complex.columns, index=None)
        if object_name_column and complex_name_column:
            df_object_complex = df_object_complex[[object_name_column, complex_name_column]]
            df_object_complex.set_index([object_name_column], inplace=True)

    if uploaded_file is not None and file_object_complex is not None:
        if object_name_column is None or complex_name_column is None:
            st.error("Zorg ervoor dat je zowel de kolom met objectnamen als de kolom met complexnamen hebt geselecteerd.")
        else:
            # Show success message
            st.success("Beide bestanden zijn succesvol geüpload! Je kunt nu de stappen rechts volgen om de conversie uit te voeren.")
            st.session_state['complex_mapping'] = df_object_complex

#On the main page, give user the option to preview both uploaded files that can be hidden/collapsed by clicking on an arrow
if 'df' in st.session_state:
    with st.expander("📊 Bekijk de OMS-export data", expanded=False):
        st.dataframe(st.session_state['df'])
if 'complex_mapping' in st.session_state:
    with st.expander("📊 Bekijk de object-complex mapping data", expanded=False):
        st.dataframe(st.session_state['complex_mapping'])

if 'df' in st.session_state and 'complex_mapping' in st.session_state:
    st.markdown("**Start de conversie naar OHJP:**")
    # Ask user for start year and start week for the planning
    start_year = st.number_input("Kies het startjaar voor de planning:", min_value=2025, max_value=2100)
    start_week = st.number_input("Kies de startweek voor de planning:", min_value=1, max_value=52, value=36)
    naam_export = st.text_input("Kies de naam voor het exportbestand (zonder extensie):", value="OHJP [X]e contractjaar [PROJECT]")
    version_export = st.pills("Kies de versie van de OHJP-export:", ["1: Definitieve versie", "2: Tijdelijke versie met kolommen 'Nummer', 'Taakplannr.' en 'Route'"], default="1: Definitieve versie")

    if st.button("Start conversie"):
        with st.spinner("Bezig met het converteren van de OMS-export naar OHJP...", show_time=True):
            df = st.session_state['df']
            df = filter_columns(df)
            df = df.dropna(subset=["Omschrijving"]) #If Omschrijving is empty, drop row
            df = df.apply(normalize_frequency, axis=1) # Transform all frequencies that are not in Months to Months
            df = df.apply(extract_uitvoerende, axis=1) # Extract executing party from Omschrijving
            df['Complex'] = df['Object'].map(df_object_complex[complex_name_column])
            #Turn Startdatum and Einddatum wk into integers
            df["Startdatum wk"] = df["Startdatum wk"].astype(int)
            df["Einddatum wk"] = df["Einddatum wk"].astype(int)
            #Extract week number from last two digits of Startdatum wk
            df["Week"] = df["Startdatum wk"].apply(lambda x: int(str(x)[-2:]))
            df["Week_end"] = df["Einddatum wk"].apply(lambda x: int(str(x)[-2:]))
            df['Weeks'] = df.apply(lambda row: list(range(row['Week'], row['Week_end'] + 1)), axis=1) # To do: Implement this in create_heatmap_df
            #Check if df["Complex"] has no empty values
            if df["Complex"].isnull().any():
                st.write("Er zijn objecten in de OMS-export die niet overeenkomen met de complexen in de mapping. Hier is een lijst van objecten die niet overeenkomen met een complex:")
                missing_complexes = df[df["Complex"].isnull()]["Object"].unique()
                st.write(missing_complexes)
                st.warning("Zorg ervoor dat de objecten in de mapping staan. De bovengenoemde objecten worden nu onder complex 'Overig' geplaatst.")
                df = df.fillna({"Complex": "Overig"})  # Fill missing complexes with 'Overig'
            # If there are values >= "start_year+1"+"start_week", then print a warning and remove those values
            week_threshold = str(start_year + 1) + str(start_week)
            df_to_remove = df[df["Startdatum wk"].astype(str) >= week_threshold]
            if not df_to_remove.empty:
                st.warning(f"Er zijn taken gepland na week {start_week} van {start_year + 1}. Deze worden niet meegenomen in de planning.")
            df = df[df["Startdatum wk"].astype(str) < week_threshold] # Remove rows

            # TEMP: Add column 'Route' with empty values
            df['Route'] = ''

            #Create Heatmap df's per object.
            complexes = df['Complex'].unique()
            heatmap_dfs_complex = {}

            for complex in complexes:
                # Filter the dataframe for the current traject
                filtered_df = df[df['Complex'] == complex]
                complex_df = create_heatmap_df(filtered_df,start_week=start_week)
                complex_df = adapt_to_version(complex_df, version=1 if version_export == "1: Definitieve versie" else 2)
                heatmap_dfs_complex[complex] = complex_df

            # Save the heatmap dataframes to separate Excel worksheets with the name of the traject
            with pd.ExcelWriter(f'{naam_export}.xlsx') as writer:
                for complex, heatmap_df in heatmap_dfs_complex.items():
                    heatmap_df = heatmap_df.drop(columns="Complex")
                    heatmap_df.to_excel(writer, sheet_name=complex, index=False)
        st.success(f"Conversie voltooid! De OHJP-gegevens worden opgeslagen als '{naam_export}.xlsx'.")
        st.write("Je kunt het bestand downloaden via de onderstaande knop:")
        st.download_button(
            label="Download OHJP Excel-bestand",
            data=open(f'{naam_export}.xlsx', 'rb').read(),
            file_name=f'{naam_export}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.write("Kopieer en plak de inhoud van de Excel per worksheet naar een OHJP-template dat past bij jouw wensen.")
        st.balloons()  # Show balloons to celebrate the successful conversion