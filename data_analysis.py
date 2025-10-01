import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.svm import LinearSVR
from sklearn.metrics import mean_squared_error
from sklearn.impute import SimpleImputer
import matplotlib.pyplot as plt

# Load the Excel file
df = pd.read_excel('Cagliari-Def.xlsx')
df = df.drop(['Link','Data rilevamento','Città','CAP','Via','Civico','Anno di costruzione'],axis=1)

si_no_columns = df.columns[df.isin(['Si', 'No']).any()]
df[si_no_columns] = df[si_no_columns].replace({'Si': 1, 'No': 0})

si_no_columns = df.columns[df.isin(['si', 'no']).any()]
df[si_no_columns] = df[si_no_columns].replace({'si': 1, 'no': 0})



# Create a new boolean column for 'Immobile a reddito'
df['Immobile_a_reddito'] = df['Contratto'].apply(lambda x: True if 'Immobile a reddito' in str(x) else False)
# Drop the original 'Contratto' column if no longer needed
df.drop(columns=['Contratto'], inplace=True)

# Convert 'disponibilità' to boolean
df['Disponibilità'] = df['Disponibilità'].apply(lambda x: True if pd.isnull(x) or 'Libero' in str(x) else False)

# Remove trailing dots from 'Efficienza energetica' values
df['Efficienza energetica'] = df['Efficienza energetica'].str.rstrip('.')
# Convert 'Efficienza energetica' to float, handling non-numeric values
df['Efficienza energetica'] = pd.to_numeric(df['Efficienza energetica'], errors='coerce')

# Remove non-numeric 'Prezzo' rows and convert to float
df = df[df['Prezzo'].str.contains('Prezzo su richiesta') == False]
df['Prezzo'] = df['Prezzo'].replace('[\€,]', '', regex=True).astype(float)

numerical_cols = df.select_dtypes(include=['int64', 'float64']).columns
categorical_cols = df.select_dtypes(include=['object']).columns

# Impute missing values
num_imputer = SimpleImputer(strategy='median')
cat_imputer = SimpleImputer(strategy='most_frequent')
df[numerical_cols] = num_imputer.fit_transform(df[numerical_cols])
df[categorical_cols] = cat_imputer.fit_transform(df[categorical_cols])

df = pd.get_dummies(df, columns=categorical_cols)

# Scale numerical features
scaler = StandardScaler()
df[numerical_cols] = scaler.fit_transform(df[numerical_cols])

# Separate features and target
X = df.drop(['Prezzo', 'Superficie', 'Prezzo unitario'], axis=1)
y = df['Prezzo unitario']

# Split the dataset
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Train a Linear SVM model
model = LinearSVR(random_state=42)
model.fit(X_train, y_train)

# Predict and evaluate
y_pred = model.predict(X_test)
mse = mean_squared_error(y_test, y_pred)
rmse = mse ** 0.5
print(f'RMSE: {rmse}')

# 9. Extract feature importance (coefficients)
feature_importance = pd.Series(model.coef_, index=X.columns)

# 10. Identify 'Tipologia' features
tipologia_features = feature_importance[feature_importance.index.str.contains('Tipologia_')]

# Aggregate the importance of dummy variables for each non-'Tipologia' categorical feature
feature_importance_df = feature_importance[~feature_importance.index.str.contains('Tipologia_')].reset_index()
feature_importance_df.columns = ['Feature', 'Importance']
feature_importance_df['Original Feature'] = feature_importance_df['Feature'].str.split('_').str[0]  # Adjust based on your naming convention

# Sum the absolute importances for each original feature
aggregated_importance = feature_importance_df.groupby('Original Feature')['Importance'].apply(lambda x: x.abs().sum()).sort_values(ascending=False)

# Combine the 'Tipologia' features and aggregated importance
combined_importance = pd.concat([tipologia_features.abs(), aggregated_importance])

# Convert the combined importance to percentages
combined_importance_percentage = (combined_importance / combined_importance.sum()) * 100

# 11. Plot the importances
top_features = combined_importance_percentage.sort_values(ascending=False)

plt.figure(figsize=(10, 6))
top_features.plot(kind='barh', color='skyblue')
plt.xlabel('Importance (%)')
plt.title('Feature Importance in Linear SVM (as Percentages)')
plt.gca().invert_yaxis()  # To have the highest value at the top

# Add annotations for clarity
for index, value in enumerate(top_features):
    plt.text(value + 1, index, f"{value:.2f}%", va='center')

# Add more ticks on the x-axis
plt.xticks(ticks=range(0, 101, 10))  # Adding ticks from 0 to 100 with a step of 10

# Show the plot
plt.show()