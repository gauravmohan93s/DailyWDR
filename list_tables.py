
import pandas as pd
from sqlalchemy import create_engine, text
from urllib.parse import quote_plus
import warnings

warnings.filterwarnings("ignore")

CONFIG = {
    "DB_SERVER": "sql-web-edition.cqu5cuepe15y.ap-south-1.rds.amazonaws.com",
    "DB_NAME":   "Kc_WebAppDb_New_Migrated",
    "ODBC_DRIVER": "ODBC Driver 17 for SQL Server",
    "DB_UID": "gaurav_sakhare_prod",
    "DB_PWD": "{G>865yQhe5NP",
}

def get_source_db_engine():
    uid, pwd = CONFIG["DB_UID"], CONFIG["DB_PWD"]
    conn_str = f"mssql+pyodbc://{quote_plus(uid)}:{quote_plus(pwd)}@{CONFIG['DB_SERVER']}/{CONFIG['DB_NAME']}?driver={quote_plus(CONFIG['ODBC_DRIVER'])}&Encrypt=yes&TrustServerCertificate=yes"
    return create_engine(conn_str)

def list_tables():
    engine = get_source_db_engine()
    query = text("SELECT table_name FROM information_schema.tables WHERE table_schema = 'dbo' ORDER BY table_name")
    with engine.connect() as conn:
        df = pd.read_sql(query, conn)
    print(df.to_string())

if __name__ == "__main__":
    list_tables()
