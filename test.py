import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image
from reportlab.lib.units import cm
import pandas as pd
import os

def create_pdf(dataframe, logos, output_path):
    """Générer un PDF avec un tableau des notes et des logos."""
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    elements = []

    # Ajouter les logos côte à côte
    logo1_img = Image(logos[0], width=4 * cm, height=1.5 * cm)
    logo2_img = Image(logos[1], width=4 * cm, height=1.5 * cm)
    logo_table = Table([[logo1_img, logo2_img]], colWidths=[5 * cm, 5 * cm])
    elements.append(logo_table)
    elements.append(Spacer(1, 1 * cm))

    # Préparer les données du tableau
    table_data = [dataframe.columns.tolist()] + dataframe.values.tolist()

    # Créer le tableau des notes
    table = Table(table_data, repeatRows=1)

    # Style du tableau
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)
    elements.append(table)

    # Générer le PDF
    doc.build(elements)

# Interface Streamlit
st.title("Générateur de relevés de notes")

st.write("Téléversez votre tableau Excel de note.")

uploaded_excel = st.file_uploader("Téléversez votre fichier Excel", type=["xls", "xlsx"])

if uploaded_excel is not None:
    df = pd.read_excel(uploaded_excel)

    # Remplacer NaN par des chaînes vides
    df.fillna('', inplace=True)

    # Nettoyer les colonnes nommées 'Unnamed'
    cleaned_columns = [col if not col.startswith("Unnamed") else "" for col in df.columns]
    df.columns = cleaned_columns

    # Chemins des logos
    logo_1 = "logo-tps-unistra-court.jpg"
    logo_2 = "logo-itii-alsace.png"

    if not os.path.exists(logo_1) or not os.path.exists(logo_2):
        st.error("Les fichiers des logos sont introuvables. Assurez-vous qu'ils se trouvent dans le même dossier que ce script.")
    else:
        output_pdf = "student_grades.pdf"
        create_pdf(df, [logo_1, logo_2], output_pdf)

        with open(output_pdf, "rb") as pdf_file:
            st.download_button(
                label="Télécharger le PDF généré",
                data=pdf_file,
                file_name="releve_notes.pdf",
                mime="application/pdf"
            )

        st.success("PDF généré avec succès !")
