from docx import Document
from docx2pdf import convert
from datetime import datetime
import locale
import polars as pl


def generate_certificate_pdf(template_path, output_pdf_path, data):
    """
    Generates a PDF certificate from a DOCX template, replacing placeholders with provided data.

    Args:
        template_path (str): Path to the DOCX template file.
        output_pdf_path (str): Path to save the generated PDF file.
        data (dict): Dictionary containing data to replace placeholders in the template.
                       Keys should correspond to the placeholder names in the DOCX template.
    """

    try:
        # 1. Load the DOCX template
        document = Document(template_path)

        # 2. Replace placeholders in paragraphs
        for paragraph in document.paragraphs:
            for placeholder_name, placeholder_value in data.items():
                if (
                    f"{{{{ {placeholder_name} }}}}" in paragraph.text
                ):  # Using Jinja-like syntax for placeholders
                    paragraph.text = paragraph.text.replace(
                        f"{{{{ {placeholder_name} }}}}", str(placeholder_value)
                    )

        # 3. Save the modified DOCX file (optional, but good for debugging or keeping a record)
        temp_docx_path = "temp_certificate.docx"  # Temporary file to save modified docx
        try:
            document.save(temp_docx_path)
            print(f"Modified DOCX template saved to '{temp_docx_path}'")
        except Exception as e:
            print(f"Error saving modified DOCX template: {e}")

        # 4. Convert the modified DOCX to PDF
        convert(temp_docx_path, output_pdf_path)
        print(f"Successfully generated PDF certificate: '{output_pdf_path}'")

    except Exception as e:
        print(f"Error generating certificate: {e}")


if __name__ == "__main__":
    # --- Example Usage ---
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    # first argv is the template path
    # second argv is the output path
    template_docx_path = "input\Declaração de Créditos Complementares - Template.docx"  # sys.argv[1]
    horas_semanais = 4
    ano = 2024
    diretor_academico = "Sara Vitória Vale Ferreira"
    df = pl.read_csv("data/Certificado CEUE (Respostas) - Form Responses 1.csv")
    
    # Loop through each row in the dataframe
    current_date = datetime.now().strftime("%d de %B de %Y")
    
    for row in df.iter_rows(named=True):
        nome_completo = row["Nome Completo"]
        mes_entrada = row["Mês que entrou no CEUE"]
        mes_saida = row["Mês que saiu do CEUE"]
        
        # Calculate months only once
        try:
            mes_saida = int(mes_saida)
        except Exception as e:
            mes_saida = 0
        meses_totais = 13 - mes_entrada - mes_saida
        horas_totais = horas_semanais * 4 * meses_totais
        
        # Determine gender-specific articles once
        is_male = row["G"] == "H"
        artigo_min = "o" if is_male else "a"
        artigo_mai = artigo_min.upper()
        
        certificate_data = {
            "o/a": artigo_min,
            "O/A": artigo_mai,
            "estudante": nome_completo.title(),
            "curso": row["Curso"],
            "cartao": str(row["Cartão UFRGS"]).zfill(8),
            "setor": row["Setor do CEUE "],
            "horas_semanais": horas_semanais,
            "meses_total": meses_totais,
            "horas_totais": horas_totais,
            "ano": ano,
            "data_documento": current_date,
            "diretor": diretor_academico,
        }
        
        output_pdf_file_path = f"output/Declaração de Créditos Complementares - {nome_completo}.pdf"
        generate_certificate_pdf(template_docx_path, output_pdf_file_path, certificate_data)
