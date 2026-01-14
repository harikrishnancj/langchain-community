from langchain_core.documents import Document
import pandas as pd

class ExcelHLoader:
    """
    Load an Excel file and convert each row to a Document.

    Handles numbers, dates, missing data, and creates one Document per row.
    """

    def __init__(self, file_path, columns=None):
        """
        file_path: path to Excel file
        columns: list of columns to include (optional)
        """
        self.file_path = file_path
        self.columns = columns

    def load(self):
        df = pd.read_excel(self.file_path)

        if self.columns:
            df = df[self.columns]

        documents = []
        for idx, row in df.iterrows():
            parts = []
            for col, val in row.items():
                if pd.isna(val):
                    val = ""
                elif isinstance(val, float):
                    val = f"{val:.2f}"
                elif isinstance(val, pd.Timestamp):
                    val = val.strftime("%Y-%m-%d")
                parts.append(f"{col}: {val}")
            documents.append(
                Document(page_content=" | ".join(parts), metadata={"row": idx})
            )
        return documents
