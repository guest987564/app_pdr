"""
Streamlit Excel Editor Application
=================================

This Streamlit app allows users to upload an Excel file (`.xlsx`), choose a
worksheet to edit, make changes directly in the browser, and then download the
updated file.  It tries to use the `streamlit-aggrid` component for a
feature‑rich editable grid.  If that package isn't available, it falls back
to Streamlit's built‑in editable table.

Key features
------------

* Upload any `.xlsx` file and select the worksheet you want to work on.
* Edit cell values inline; sort and filter when AgGrid is available.
* Download the entire workbook with your changes applied while preserving
  untouched worksheets.
* Runs entirely in the browser — ideal for Streamlit Cloud deployments.
"""

import io
from typing import List

import pandas as pd
import streamlit as st


def try_import_aggrid():
    """Try importing AgGrid; return a tuple (available, AgGrid class, options helpers)."""
    try:
        from st_aggrid import AgGrid
        from st_aggrid import GridOptionsBuilder
        from st_aggrid import GridUpdateMode
        from st_aggrid import DataReturnMode
        return True, AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode
    except Exception:
        return False, None, None, None, None


def load_excel(file) -> pd.ExcelFile:
    """Load an uploaded file into a pandas ExcelFile object."""
    try:
        return pd.ExcelFile(file)
    except Exception as exc:
        st.error(f"Failed to read Excel file: {exc}")
        raise


def build_aggrid_table(df: pd.DataFrame, AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode) -> pd.DataFrame:
    """Render an editable AgGrid and return the edited dataframe."""
    builder = GridOptionsBuilder.from_dataframe(df)
    # Make all columns editable by default.  Users can still sort/filter.
    builder.configure_default_column(editable=True)
    grid_options = builder.build()
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.VALUE_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        allow_unsafe_jscode=True,
        theme="light",
        fit_columns_on_grid_load=True,
    )
    return pd.DataFrame(grid_response["data"])


def build_data_editor_table(df: pd.DataFrame) -> pd.DataFrame:
    """Render Streamlit's editable table and return the edited dataframe."""
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        key="data_editor",
    )
    return edited_df


def write_workbook(edits: pd.DataFrame, sheet_name: str, xl: pd.ExcelFile) -> bytes:
    """Create a new Excel workbook with the edited sheet and return its bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Write the edited sheet first
        edits.to_excel(writer, index=False, sheet_name=sheet_name)
        # Copy other sheets unchanged
        for name in xl.sheet_names:
            if name == sheet_name:
                continue
            try:
                df_other = xl.parse(name)
            except Exception:
                df_other = pd.DataFrame()
            df_other.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()


def main() -> None:
    st.set_page_config(page_title="Excel Editor App", layout="wide")
    st.title("📊 Internal Excel Editor")

    st.write(
        """
        Téléversez un fichier Excel (`.xlsx`), choisissez la feuille à modifier puis
        éditez les données directement dans le tableau ci‑dessous.
        """
    )

    aggrid_available, AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode = try_import_aggrid()

    st.sidebar.title("📁 Options")
    st.sidebar.write(
        "1. Téléversez un fichier\n2. Choisissez la feuille\n3. Téléchargez le classeur"
    )

    uploaded_file = st.sidebar.file_uploader("Choisir un fichier Excel", type=["xlsx"])
    if not uploaded_file:
        st.info("Aucun fichier sélectionné. Veuillez sélectionner un fichier `.xlsx`.")
        return

    # Load the workbook
    xl = load_excel(uploaded_file)
    sheets: List[str] = xl.sheet_names
    if not sheets:
        st.error("Ce classeur ne contient aucune feuille.")
        return

    # Choose a sheet
    sheet_name = st.sidebar.selectbox("Sélectionner la feuille à éditer", options=sheets)
    if not sheet_name:
        return

    # Parse selected sheet into a dataframe
    try:
        df = xl.parse(sheet_name)
    except Exception as exc:
        st.error(f"Impossible de lire la feuille '{sheet_name}': {exc}")
        return

    # Display some quick metrics
    col1, col2 = st.columns(2)
    col1.metric("Lignes", df.shape[0])
    col2.metric("Colonnes", df.shape[1])

    # Display editable table using AgGrid if available
    st.subheader(f"Édition de la feuille : {sheet_name}")
    if aggrid_available:
        st.caption("Vous pouvez trier, filtrer et éditer les cellules directement dans la grille.")
        edited_df = build_aggrid_table(df, AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode)
    else:
        st.caption("Module AgGrid non disponible ; utilisation de l'éditeur Streamlit intégré.")
        edited_df = build_data_editor_table(df)

    # Provide download button in the sidebar
    st.sidebar.markdown("---")
    data_bytes = write_workbook(edited_df, sheet_name, xl)
    st.sidebar.download_button(
        label="💾 Télécharger le classeur modifié",
        data=data_bytes,
        file_name=f"modifié_{uploaded_file.name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
