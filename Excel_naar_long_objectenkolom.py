"""
Script om Excel bestanden in een map te verwerken en om te zetten naar een lang formaat voor PowerBI
Vereisten:
- Python 3.7+
- pandas, numpy
Gebruik:
python Excel_naar_long.py
"""

import pandas as pd
import numpy as np
import os

# Pad naar map met Excel-bestanden (aanpassen!)
pad_naar_excelmap = r"C:\Users\bjmrijve\OneDrive - Avans Hogeschool\Documents\Ziekenhuis Microbioom\full workflow\demo\infectie_data"

def lees_datum_uit_cel(xls, sheet, bestand):
    """
    Leest een datum uit cel C2 van een opgegeven Excel-blad en retourneert het jaar en de maand.

    Parameters:
        xls (str or file-like object): Het pad naar het Excel-bestand of een bestand-object.
        sheet (str): De naam van het werkblad waaruit de datum gelezen moet worden.
        bestand (str): De naam van het bestand (voor foutmeldingen).

    Returns:
        tuple: Een tuple met het jaar (int) en de maand (int) van de gelezen datum.

    Raises:
        Exception: Als er geen datum gevonden wordt in cel C2 of als er een fout optreedt bij het lezen van de datum.
    """
    try:
        # Lees datum uit cel C2
        datumcel = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=2, usecols="C")
        datumwaarde = datumcel.iloc[1, 0]
        if pd.isna(datumwaarde):
            # Error als er geen datum is
            raise Exception(f"Geen datum gevonden in C2: {bestand}, tabblad: {sheet}")
        datum = pd.to_datetime(datumwaarde, format='%d-%m-%Y')
        return datum.year, datum.month
    # Error als geen datum kan worden omgezet
    except Exception as e:
        raise Exception(f"Fout bij het lezen van datum uit C2: {bestand}, tabblad: {sheet}, foutmelding: {e}")

def verwerk_sheet(xls, sheet, bestand, jaar, maand):
    """
        Verwerkt een Excel-tabblad en zet het om naar een long-format DataFrame.
        Parameters:
            xls (str of file-like object): Pad naar het Excel-bestand of een bestand-object.
            sheet (str of int): Naam of index van het tabblad dat verwerkt moet worden.
            bestand (str): Naam van het bestand dat wordt verwerkt (voor foutmeldingen).
            jaar (int): Jaar dat aan de data is gekoppeld.
            maand (int): Maand die aan de data is gekoppeld.
        Returns:
            pandas.DataFrame: Een long-format DataFrame met de volgende kolommen:
                - 'Organisme': De naam van het organisme.
                - 'Resistentie': Informatie over resistentie.
                - 'Afdeling': Naam van de afdeling.
                - 'Waarde': Waarde gekoppeld aan de afdeling.
                - 'BRMO': Binaire indicator voor BRMO-gerelateerde resistentie.
                - 'ESBL': Binaire indicator voor ESBL-gerelateerde resistentie.
                - 'VRE': Binaire indicator voor VRE-gerelateerde resistentie.
                - 'MRSA': Binaire indicator voor MRSA-gerelateerde resistentie.
                - 'CARBA': Binaire indicator voor CARBA-gerelateerde resistentie.
                - 'Jaar': Jaar van de data.
                - 'Maand': Maand van de data.
        Raises:
            Exception: Als het tabblad leeg is of als er een fout optreedt tijdens de verwerking.
                    Het foutbericht bevat de naam van het bestand, de naam van het tabblad en details over de fout.
        Notes:
            - De functie slaat de eerste twee rijen van het tabblad over en verwerkt de data vanaf de derde kolom.
            - Lege rijen en kolommen worden verwijderd.
            - Non-breaking spaces in stringkolommen worden vervangen door gewone spaties.
            - Dubbele kolommen worden verwijderd.
            - De data wordt omgevormd naar een long-format met behulp van de 'melt'-functie.
            - Kolommen gerelateerd aan resistentie (bijv. 'BRMO', 'ESBL', enz.) worden afgeleid met behulp van stringmatching.
    """
    try:
        # Lees tabel vanaf rij 3 en kolom C (index 2)
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=2).iloc[:, 2:]
        # Verwijder lege rijen en kolommen
        df = df.dropna(how='all').dropna(axis=1, how='all')
        df.columns = df.columns.str.strip()
        # Non breaking spaces normaliseren in de dataframe
        df = df.map(lambda x: x.replace('\xa0', ' ') if isinstance(x, str) else x)

        # Maak een nieuwe kolom voor alleen organisme namen uit kolom C (eerste kolom df, index 0)
        df['Organisme'] = df.iloc[:, 0].astype(str).replace('nan', '').str.strip().replace('', np.nan)
        # Maak een nieuwe kolom voor alleen resistentie uit kolom D (tweede kolom df, index 1)
        df['Resistentie'] = df.iloc[:, 1].astype(str).replace('nan', '').str.strip().replace('', np.nan)

        afdelingskolommen = df.columns[2:-1] if df.columns[-1] == 'Organisme' else df.columns[2:]
        df = df[['Organisme', 'Resistentie'] + list(afdelingskolommen)]
        # Verwijder dubbele kolommen
        df = df.loc[:, ~df.columns.duplicated()]

        # Zet in long format
        long_df = df.melt(
            id_vars=['Organisme', 'Resistentie'],
            var_name='Afdeling',
            value_name='Waarde'
        )
        
        # Error als er geen data is
        if long_df.empty:
            raise Exception(f"Geen data gevonden in tabel: {bestand}, tabblad: {sheet}")
        
        # Binaire kolommen maken voor resistentie
        long_df['BRMO'] = long_df['Resistentie'].str.contains(r'BRMO|ESBL|VRE|MRSA|CARBA', case=False, na=False).astype(int)
        long_df['ESBL'] = long_df['Resistentie'].str.contains('ESBL', case=False, na=False).astype(int)
        long_df['VRE'] = long_df['Resistentie'].str.contains('VRE', case=False, na=False).astype(int)
        long_df['MRSA'] = long_df['Resistentie'].str.contains('MRSA', case=False, na=False).astype(int)
        long_df['CARBA'] = long_df['Resistentie'].str.contains('CARBA', case=False, na=False).astype(int)
        
        # Datum gegevens toevoegen
        long_df['Jaar'] = jaar
        long_df['Maand'] = maand

        return long_df
    except Exception as e:
        raise Exception(f"Fout bij het inlezen van tabel: {bestand}, tabblad: {sheet}, foutmelding: {e}")

def verwerk_excelbestand(bestandpad, bestand):
    """
    Verwerkt een Excel-bestand door elk werkblad te lezen en te verwerken.
    Parameters:
        bestandpad (str): Het volledige pad naar het Excel-bestand.
        bestand (str): De naam van het Excel-bestand.
    Returns:
        list: Een lijst van DataFrames, waarbij elk DataFrame de verwerkte gegevens
              van een werkblad in het Excel-bestand bevat.
    Raises:
        Exception: Als het bestand niet geopend kan worden of als er een fout optreedt
                   tijdens het verwerken van het bestand.
    """
    try:
        xls = pd.ExcelFile(bestandpad)
    except Exception as e:
        raise Exception(f"Fout bij het openen van bestand: {bestand}, foutmelding: {e}")

    resultaat_per_bestand = []
    for sheet in xls.sheet_names:
        jaar, maand = lees_datum_uit_cel(xls, sheet, bestand)
        long_df = verwerk_sheet(xls, sheet, bestand, jaar, maand)
        resultaat_per_bestand.append(long_df)

    return resultaat_per_bestand

def filter_exclusie(df):
    """
    Filtert een DataFrame om alleen de rijen te behouden die niet onder de exclusiecriteria vallen.
    Parameters:
        df (pandas.DataFrame): De DataFrame die gefilterd moet worden.
    Returns:
        df_behouden (pandas.DataFrame): De DataFrame met alleen de rijen die niet onder de exclusiecriteria vallen.
        df_exclusie (pandas.DataFrame): De DataFrame met alleen de rijen die onder de exclusiecriteria vallen.
    """
    verboden_woorden = ['coccen', 
                        'staven', 
                        'Coagulase-negatieve', 
                        'Fluoresc. preparaat', 
                        'vergroenend', 
                        'gekweekt',
                        'BRMO PCR',
                        'ESBL/CPE PCR', 
                        'IgG',
                        'IgM'
                        ]
    verboden_taxonomie = ['Plasmodium', 
                          'C.psittaci', 
                          'Q-koorts', 
                          'Rotavirus antigeen sneltest',
                          'Adenovirus antigeen sneltest',
                          'MRSA PCR sneltest',
                          'MRSA PCR',
                          'MRSA DNA',
                          'VRE PCR sneltest'
                          ]
    # Filteren op verboden woorden
    mask_verboden_woorden = df['Organisme'].str.contains('|'.join(verboden_woorden), case=False, na=False)
    # Filteren op verboden taxonomie
    mask_verboden_taxonomie = df['Organisme'].str.contains('|'.join(verboden_taxonomie), case=False, na=False)
    # Filteren op komma
    mask_komma = df['Organisme'].str.contains(',', na=False)

    # Alle maskers combineren
    mask_exclusie = (mask_verboden_woorden | mask_verboden_taxonomie | mask_komma)
    
    # Schrijf nieuwe dataframes
    # Met alleen behouden data
    df_behouden = df[~mask_exclusie]
    # Met alleen verwijderde data
    df_exclusie = df[mask_exclusie]
    
    return df_behouden, df_exclusie

def main():
    alle_data = []
    for bestand in os.listdir(pad_naar_excelmap):
        if bestand.lower().endswith(('.xlsx', '.xlsm')):
            bestandpad = os.path.join(pad_naar_excelmap, bestand)
            alle_data.extend(verwerk_excelbestand(bestandpad, bestand))

    if not alle_data:
        raise Exception("Geen data gevonden in de opgegeven map.")

    try:
        res_binair = pd.concat(alle_data, ignore_index=True)
        res_binair_no_0 = res_binair[res_binair['Waarde'] != 0]

        # Filteren op exclusie criteria
        resultaat, exclusie = filter_exclusie(res_binair_no_o)

        # Zorgen dat PowerBI de juiste variabelen kan vinden
        globals().update(locals())
    except Exception as e:
        raise Exception(f"Fout bij het combineren van dataframes, foutmelding: {e}")

# Main aanroepen als dit script direct wordt uitgevoerd
main()
