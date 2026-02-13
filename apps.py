import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Transformer",
                   page_icon="📊",
                   layout="centered")
st.title("📊 Jen GSL Report Transformer")
st.write("A fix for LT Lam and all her friends for that notorious GSL report.")


def gslreports(filename):
    df = pd.read_excel(filename)

    # Validate minimum data requirements
    if len(df) <= 10:
        raise ValueError("File must have more than 10 rows of data")
    if len(df.columns) < 3:
        raise ValueError("File must have at least 3 columns")

    cols = df.columns

    # input data processing
    df = df.iloc[10:]  # Skip header rows
    # Keep first, second-to-last, and last columns
    df = df[[cols[0], cols[-2], cols[-1]]]
    df.ffill(inplace=True)  # Fill NaN values from merged cells
    output = pd.DataFrame(columns=['Name', 'Rx', 'Baskets'])

    names = list(df[df.columns[0]].unique())

    for name in names:
        output.loc[len(output.index)] = \
            df[df[cols[0]] == name].nlargest(1, cols[-2]).iloc[0].tolist()

    return output


uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

if uploaded:
    try:
        df = gslreports(uploaded)

        # Save to Excel in-memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Transformed")
        output.seek(0)

        # Generate output filename based on original filename
        output_filename = uploaded.name.split('.')[0] + ' Processed.xlsx'

        mime_type = (
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
        st.download_button(
            label="⬇️ Download transformed GSL report",
            data=output,
            file_name=output_filename,
            mime=mime_type,
        )
        st.success("✅ File processed successfully!")
    except ValueError as e:
        st.error(f"Data validation error: {e}")
    except Exception as e:
        st.error(f"Unexpected error processing file: {e}")
