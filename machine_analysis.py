import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.svm import SVR
from sklearn.metrics import mean_absolute_error

df = pd.read_excel("Cagliari-Def.updated.xlsx", sheet_name=0, engine="openpyxl")

# example schema
cat_cols = ["Tipologia", "Quartiere", "Riscaldamento"]
num_cols = ["Superficie_mq", "AnnoCostruzione", "Piano"]

X = df.drop("price", axis=1)
y = df["price"]
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# categorical: fill with explicit token and add missing indicator
cat_imputer = SimpleImputer(strategy="constant", fill_value="__MISSING__")
ohe = OneHotEncoder(handle_unknown="ignore", sparse_output=True)

cat_pipe = Pipeline([
    ("impute", cat_imputer),
    ("encode", ohe),
])

# numeric: median + missing flag
num_imputer = SimpleImputer(strategy="median", add_indicator=True)

pre = ColumnTransformer([
    ("cat", cat_pipe, cat_cols),
    ("num", num_imputer, num_cols),
])

svm = SVR(kernel="rbf", C=10.0, gamma="scale", epsilon=0.1)

model = Pipeline([
    ("pre", pre),
    ("svm", svm),
])

model.fit(X_train, y_train)
pred = model.predict(X_test)
print("MAE:", mean_absolute_error(y_test, pred))
