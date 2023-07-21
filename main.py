import pandas as pd
import streamlit as st
from streamlit import components
import os
from io import StringIO

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


def excel_file_opener(file):
    df = pd.read_excel(file)
    return df

def df_cleaner(df):

    for col in df.columns:

        if col not in ["VOLUME", "JO"]:
            df[col] = df[col].fillna(method='ffill')
        else:
            df[col] = df[col].fillna(0)
    df.drop(df[df["VOLUME"] == 0].index, inplace=True)
    pass


def set_page(st):
    st.set_page_config(page_title="Excel Comparator",
                       page_icon="❄️",
                       initial_sidebar_state="auto",layout="wide")
    st.header("📖Excel Comparator")
    pass

def sidebar():
    with st.sidebar:
        st.markdown(
            "# Comment l'utiliser\n"
        )
        st.markdown("---")
        st.markdown("## Pour insérer les excels:")
        st.markdown("### Télécharger les 2 excels consécutivement")
        st.markdown("---")
        st.markdown("## Pour réinitialiser:")
        st.markdown("### Veuillez cliquer sur le bouton actualiser du navigateur")
        st.markdown("---")
        st.markdown("## Informations:")
        st.markdown("### Pour l'instant, on ne peut pas déposer de csv")
        st.markdown("---")
        st.markdown("## Traitements:")
        st.markdown("### Le fichier en entrée est nettoyé des lignes VOLUME = 0 et les cases fusionnées sont séparées")
        st.markdown("---")
    pass


def main():

    set_page(st)
    #sidebar()

    my_expander = st.expander(label='Télécharger les fichiers ici', expanded=True)
    colA,colB = my_expander.columns(2)
    colA.header("Original")
    colB.header("Généré")
    original = colA.file_uploader("Télécharger l'excel original", type=["xlsx"])
    genere = colB.file_uploader("Télécharger l'excel généré", type=["xlsx"])

    col1, col2,col3 = st.columns(3)
    col1.header("Original")
    col3.header("Généré")
    col2.header("Comparaison")


    if original and genere:
        df_original = excel_file_opener(original)
        df_cleaner(df_original)
        df_genere = excel_file_opener(genere)
        df_cleaner(df_genere)

        #analyse dimensions
        expander_dimension = st.expander(label='Dimensions', expanded=True)
        col1_dimension,col2_dimension,col3_dimension = expander_dimension.columns(3)
        col1_dimension.header("Original")
        col2_dimension.header("Diff")
        col3_dimension.header("Genere")
        col1_dimension.write("lignes: {} et colonnes {}".format(str(df_original.shape[0]), str(df_original.shape[1])))
        col3_dimension.write("lignes: {} et colonnes {}".format(str(df_genere.shape[0]), str(df_genere.shape[1])))
        col2_dimension.write("lignes: {} et colonnes {}".format(str(df_original.shape[0]-df_genere.shape[0]), str(df_original.shape[1]-df_genere.shape[1])))

        #analyse colonnes
        expander_colonne= st.expander(label='Colonnes', expanded=True)
        col1_colonne,col2_colonne,col3_colonne = expander_colonne.columns(3)
        col1_colonne.header("Original")
        col2_colonne.header("Diff")
        col3_colonne.header("Genere")
        col_orignial = df_original.columns
        col_generated = df_genere.columns
        col1_colonne.dataframe(col_orignial)
        col3_colonne.dataframe(col_generated)
        if len(col_orignial) != len(col_generated):
            col2_colonne.warning("Il n'y a pa le même nombre de colonnes")
        elif set(col_orignial) != set(col_generated):
            col2_colonne.warning("Il n'y a pa les mêmes colonnes")
        else:
            col_gauche = pd.DataFrame(col_orignial,columns=["gauche"])
            col_droite = pd.DataFrame(col_generated,columns=["droite"])
            result = pd.concat([col_gauche, col_droite], axis=1)
            result["similarity"] = result["droite"] == result["gauche"]
            def color_coding(row):
                return ['background-color:red'] * len(
                    row) if not row.similarity else ['background-color:green'] * len(row)
            result.drop(columns=["gauche"],inplace=True)
            col2_colonne.dataframe(result.style.apply(color_coding, axis=1))
        #analyse volumes et dates
        expander_volume = st.expander(label='Volumes par date', expanded=True)
        col1_volume,col2_volume,col3_volume = expander_volume.columns(3)
        col1_volume.header("Original")
        col2_volume.header("Diff")
        col3_volume.header("Genere")
        origi =pd.pivot_table(df_original, values='VOLUME', index='DATE',aggfunc='sum')
        col1_volume.dataframe(origi)
        genere =pd.pivot_table(df_genere, values='VOLUME', index='DATE',aggfunc='sum')
        col3_volume.dataframe(genere)
        if df_original["DATE"].unique().tolist() != df_genere["DATE"].unique().tolist():
            col2_volume.warning("Il n'y a pa les mêmes dates")
        else:
            result = pd.concat([origi, genere], axis=1)
            result["similarity"] = result["droite"] == result["gauche"]
            #col2_volume.dataframe(result)
            print(result)
            def color_coding(row):
                return ['background-color:red'] * len(
                    row) if not row.similarity else ['background-color:green'] * len(row)
            result.drop(columns=["gauche"],inplace=True)
            col2_volume.dataframe(result.style.apply(color_coding, axis=1))

        #analyse regroupements
        expander_regroupement = st.expander(label='Regroupements', expanded=True)
        col1_regroupement,col2_regroupement,col3_regroupement = expander_regroupement.columns(3)
        col1_regroupement.header("Original")
        col2_regroupement.header("Diff")
        col3_regroupement.header("Genere")
        #col_orignial = df_original.columns
        #col_generated = df_genere.columns
        for a in set(col_orignial).intersection(set(col_generated)):
            if a not in ["DATE","VOLUME","DATE_ARRETEE"]:
                expander=st.expander(label=a, expanded=True)
                col1_exp,col3_exp = expander.columns(2)
                left = pd.pivot_table(df_original, values='VOLUME', index=a,aggfunc='sum')
                rigt = pd.pivot_table(df_genere, values='VOLUME', index=a,aggfunc='sum')
                col1_exp.dataframe(left)
                col3_exp.dataframe(rigt)






    else:
        st.warning("Veuillez télécharger les 2 fichiers")

if __name__ == "__main__":
    main()
