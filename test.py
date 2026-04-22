import sqlite3, pandas as pd
con = sqlite3.connect('mps_plasticos.db')
df = pd.read_sql("SELECT ItemCode, Tipo_Material2 FROM bd_spec WHERE ItemCode IN ('2310360001','2310390064')", con)
print(df.to_string())
con.close()