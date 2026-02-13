import numpy as np
import pandas as pd
import warnings
from pathlib import Path
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.model_selection import GridSearchCV, KFold, RandomizedSearchCV, cross_validate, train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, OneHotEncoder, StandardScaler
from sklearn.svm import SVR
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn.neural_network import MLPRegressor

try:
    from pytorch_tabnet.tab_model import TabNetRegressor
    TABNET_AVAILABLE = True
except ImportError:
    TABNET_AVAILABLE = False

try:
    from lightgbm import LGBMRegressor
    LIGHTGBM_AVAILABLE = True
except ImportError:
    LIGHTGBM_AVAILABLE = False

try:
    import shap
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    SHAP_AVAILABLE = True
except ImportError:
    SHAP_AVAILABLE = False
    shap = None
    plt = None


if TABNET_AVAILABLE:
    class TabNetRegressorWrapper(TabNetRegressor):
        def fit(self, X, y, **kwargs):
            y_arr = np.asarray(y)
            if y_arr.ndim == 1:
                y_arr = y_arr.reshape(-1, 1)
            X_train_split, X_val, y_train_split, y_val = train_test_split(
                X, y_arr, test_size=0.1, random_state=42
            )
            X_train_split = np.asarray(X_train_split, dtype=np.float32)
            X_val = np.asarray(X_val, dtype=np.float32)
            y_train_split = np.asarray(y_train_split, dtype=np.float32)
            y_val = np.asarray(y_val, dtype=np.float32)
            kwargs.setdefault("max_epochs", 200)
            kwargs.setdefault("patience", 30)
            kwargs.setdefault("eval_metric", ["rmse"])
            kwargs.setdefault("eval_name", ["valid"])
            kwargs.setdefault("eval_set", [(X_val, y_val)])
            return super().fit(X_train_split, y_train_split, **kwargs)

        def predict(self, X):
            return super().predict(X).ravel()


float32_transformer = FunctionTransformer(
    lambda X: np.asarray(X.toarray() if hasattr(X, "toarray") else X, dtype=np.float32),
    accept_sparse=True,
)


df = pd.read_excel("Cagliari-Def.updated.xlsx", sheet_name=0, engine="openpyxl")
df = df[df["Prezzo"] != "Prezzo su richiesta"].copy()

cat_cols = [
    "Tipologia",
    "Classe Immobile",
    "Tipologia riscaldamento",
    "Tipologia di infissi",
    "Materiale infissi",
    "Posti Auto",
    "Esposizione",
    "Stato di conservazione",
    "Disponibilit\u00e0",
    "Propriet\u00e0",
    "Alimentazione riscaldamento",
    "Fonte riscaldamento",
    "Classe energetica",
]
num_cols = ["Superficie", "Anno di costruzione", "Piano", "Efficienza energetica"]

df["Prezzo"] = pd.to_numeric(df["Prezzo"], errors="coerce")
df["Superficie"] = pd.to_numeric(df["Superficie"], errors="coerce")
df = df[df["Superficie"] > 0].copy()

df["Prezzo_per_mq"] = df["Prezzo"] / df["Superficie"]

X = df.drop("Prezzo", axis=1)
y = df["Prezzo_per_mq"].values

cat_imputer = SimpleImputer(strategy="constant", fill_value="__MISSING__")
ohe = OneHotEncoder(handle_unknown="ignore", sparse_output=False)
cat_pipe = Pipeline([
    ("impute", cat_imputer),
    ("encode", ohe),
])

num_imputer = SimpleImputer(strategy="median", add_indicator=True)
pre = ColumnTransformer([
    ("cat", cat_pipe, cat_cols),
    ("num", num_imputer, num_cols),
])


def build_pipeline(regressor, *, scale=False, to_float32=False):
    steps = [("pre", pre)]
    if scale:
        steps.append(("scaler", StandardScaler()))
    if to_float32:
        steps.append(("to_float32", float32_transformer))
    steps.append(("regressor", regressor))
    return Pipeline(steps)


def describe_search_space(params):
    lines = []
    for param_name, values in params.items():
        try:
            value_list = list(values)
        except TypeError:
            value_list = [values]
        value_repr = ", ".join(str(v) for v in value_list)
        line = f"  {param_name}: [{value_repr}]"
        numeric_values = []
        numeric_only = True
        for value in value_list:
            if isinstance(value, bool):
                numeric_only = False
                break
            if isinstance(value, (int, float, np.integer, np.floating)):
                numeric_values.append(float(value))
            else:
                numeric_only = False
                break
        if numeric_only and numeric_values:
            min_val = min(numeric_values)
            max_val = max(numeric_values)
            line += f" (range: {min_val:g} to {max_val:g})"
        lines.append(line)
    return lines


def slugify(value):
    cleaned = [c.lower() if c.isalnum() else "_" for c in str(value)]
    slug = "".join(cleaned).strip("_")
    return slug or "model"

def ensure_feature_names_length(feature_names, target_len):
    names = list(feature_names or [])
    if len(names) < target_len:
        names.extend([f"feature_{i}" for i in range(len(names), target_len)])
    elif len(names) > target_len:
        names = names[:target_len]
    return names


def append_indicator_feature_names(feature_names, cols, transformer, prefix):
    """
    If a transformer added missing-value indicators, append names using the original
    column names, always as '{prefix}__{col}__missing_indicator'.
    """
    indicator = getattr(transformer, "indicator_", None)
    if indicator is None or not getattr(transformer, "add_indicator", False):
        return
    features = getattr(indicator, "features_", None)
    if features is None:
        return
    # cols can be a list of column names or indices. Normalize to names.
    if isinstance(cols, (list, tuple)):
        col_list = list(cols)
    else:
        col_list = [cols]
    for idx in features:
        if 0 <= idx < len(col_list):
            base = col_list[idx]
        else:
            base = f"unknown_col_{idx}"
        feature_names.append(f"{prefix}__{base}__missing_indicator")


def _ohe_names_with_prefix(ohe, input_cols, prefix):
    """
    Robustly get one hot names and force a stable '{prefix}__{col}_{category}'
    format even when OHE returns 'x0_' style names.
    """
    try:
        raw = list(ohe.get_feature_names_out(input_cols))
    except TypeError:
        # older sklearn
        raw = list(ohe.get_feature_names_out())
    names = []
    # Build mapping index->col for 'x0_' style
    idx_to_col = {i: c for i, c in enumerate(input_cols)}
    for j, n in enumerate(raw):
        # Cases:
        # 1) 'Tipologia_appartamento'
        # 2) 'x0_appartamento'
        # 3) 'Tipologia' if drop/handle_unknown interplay
        if n.startswith("x") and "_" in n and n[1:n.find("_")].isdigit():
            xi = int(n[1:n.find("_")])
            col = idx_to_col.get(xi, input_cols[0] if input_cols else f"col{xi}")
            cat = n[n.find("_")+1:].replace(" ", "_")
            names.append(f"{prefix}__{col}_{cat}")
        elif "_" in n:
            col = n[:n.find("_")]
            cat = n[n.find("_")+1:].replace(" ", "_")
            names.append(f"{prefix}__{col}_{cat}")
        else:
            # No category suffix returned. Keep as column name.
            names.append(f"{prefix}__{n}")
    return names


def get_pipeline_feature_names(pipeline):
    """
    Always return names with stable prefixes:
      - categorical: 'cat__{original_col}_{category}'
      - numeric:     'num__{original_col}' plus optional '__missing_indicator'
    Falls back gracefully if get_feature_names_out is not available.
    """
    preprocessor = pipeline.named_steps.get("pre") if hasattr(pipeline, "named_steps") else None
    if preprocessor is None:
        return None

    # Try the modern path first
    try:
        names = list(preprocessor.get_feature_names_out())
        if names:
            return names
    except Exception:
        pass

    # Manual reconstruction
    feature_names = []
    transformers = getattr(preprocessor, "transformers_", [])
    for name, transformer, cols in transformers:
        if transformer in ("drop", None):
            continue

        # Normalize cols to list of original column names
        if isinstance(cols, (list, tuple, np.ndarray, pd.Index)):
            input_cols = list(cols)
        else:
            input_cols = [cols]

        # Pipeline inside ColumnTransformer
        steps = None
        fitted_transformer = transformer
        if hasattr(fitted_transformer, "steps"):
            steps = fitted_transformer.steps
            fitted_transformer = steps[-1][1]

        if fitted_transformer == "passthrough":
            # Preserve original column names with the section prefix
            prefix = f"{name}"
            for c in input_cols:
                feature_names.append(f"{prefix}__{c}")
        elif hasattr(fitted_transformer, "get_feature_names_out") and isinstance(fitted_transformer, OneHotEncoder):
            # OneHotEncoder: force stable prefixed names
            ohe_names = _ohe_names_with_prefix(fitted_transformer, input_cols, name)
            feature_names.extend(ohe_names)
        elif hasattr(fitted_transformer, "get_feature_names_out"):
            # Other transformers that expose names, prefix them with the section name
            try:
                out = list(fitted_transformer.get_feature_names_out(input_cols))
            except TypeError:
                out = list(fitted_transformer.get_feature_names_out())
            for n in out:
                # Avoid double prefixing if already prefixed
                feature_names.append(n if n.startswith(f"{name}__") else f"{name}__{n}")
        else:
            # Fallback: keep original columns with section prefix
            for c in input_cols:
                feature_names.append(f"{name}__{c}")

        # Append any missing indicators from the fitted components
        prefix = name
        # indicator on the last step
        append_indicator_feature_names(feature_names, input_cols, fitted_transformer, prefix)
        # indicator possibly attached to earlier steps in the pipeline
        if steps is not None:
            for _, step_transformer in steps:
                append_indicator_feature_names(feature_names, input_cols, step_transformer, prefix)

    return feature_names or None


def base_feature_label(feature_name):
    """
    Collapse one hot levels and indicators to the base original column name.
    Works with the enforced '{sect}__{col}' naming.
    """
    name = feature_name
    if name.endswith("__missing_indicator"):
        name = name[: -len("__missing_indicator")]
    # Remove section prefix 'cat__' or 'num__'
    if "__" in name:
        parts = name.split("__", 2)
        if len(parts) >= 2:
            name = parts[1]
    # For one hot like '{col}_{category}', strip category
    if "_" in name and name.split("_", 1)[0] in set(cat_cols):
        return name.split("_", 1)[0]
    return name


def friendly_feature_label(feature_name):
    """
    Human friendly for plots:
      - 'cat__Col_cat' -> 'Col = cat'
      - 'num__Col'     -> 'Col'
      - indicators     -> 'Col (missing indicator)'
    """
    if feature_name.endswith("__missing_indicator"):
        base = friendly_feature_label(feature_name[: -len("__missing_indicator")])
        return f"{base} (missing indicator)"

    # Expect '{sect}__{payload}'
    if "__" in feature_name:
        sect, payload = feature_name.split("__", 1)
        if sect == "cat":
            # payload can be '{col}_{category}' or just '{col}'
            if "_" in payload:
                col, cat = payload.split("_", 1)
                return f"{col} = {cat.replace('_', ' ')}"
            return payload
        if sect == "num":
            return payload

    # Fallback unchanged
    return feature_name


def generate_shap_summary(pipeline, feature_names, X, model_name):
    if not SHAP_AVAILABLE:
        return None
    if feature_names is None:
        feature_names = []
    try:
        preprocess = pipeline[:-1]
    except TypeError:
        preprocess = None
    if preprocess is None:
        return {"name": model_name, "error": "Pipeline does not expose preprocessing steps for SHAP."}
    try:
        transformed = preprocess.transform(X)
    except Exception as exc:
        return {"name": model_name, "error": f"Failed to transform inputs for SHAP: {exc}"}
    if hasattr(transformed, "toarray"):
        transformed = transformed.toarray()
    transformed = np.asarray(transformed)
    if transformed.size == 0:
        return {"name": model_name, "error": "Transformed design matrix is empty."}

    dtype = transformed.dtype
    sample_size = min(SHAP_SAMPLE_SIZE, transformed.shape[0]) if SHAP_SAMPLE_SIZE else transformed.shape[0]
    rng = np.random.default_rng(SHAP_RANDOM_STATE)
    sample_idx = rng.choice(transformed.shape[0], size=sample_size, replace=False)
    X_sample = transformed[sample_idx]

    background_size = min(SHAP_BACKGROUND_SIZE, transformed.shape[0])
    if background_size == transformed.shape[0]:
        background = transformed
    else:
        background_idx = rng.choice(transformed.shape[0], size=background_size, replace=False)
        background = transformed[background_idx]

    try:
        regressor = pipeline.named_steps.get("regressor", pipeline.steps[-1][1])
    except Exception:
        regressor = None
    if regressor is None:
        return {"name": model_name, "error": "Unable to locate regressor in pipeline."}

    def predict_fn(data):
        arr = np.asarray(data, dtype=dtype)
        preds = regressor.predict(arr)
        return np.asarray(preds).reshape(-1)

    try:
        try:
            masker = shap.maskers.Independent(background)
            explainer = shap.Explainer(predict_fn, masker, algorithm="permutation")
        except AttributeError:
            explainer = shap.Explainer(predict_fn, background)
        shap_result = explainer(X_sample)
    except Exception as exc:
        return {"name": model_name, "error": f"SHAP explainer failed: {exc}"}

    shap_values = getattr(shap_result, "values", shap_result)
    data_for_plot = getattr(shap_result, "data", X_sample)
    shap_values = np.asarray(shap_values)
    data_for_plot = np.asarray(data_for_plot)
    if shap_values.ndim == 3:
        shap_values = shap_values[:, 0, :]
    if shap_values.size == 0 or shap_values.shape[1] == 0:
        return {"name": model_name, "error": "SHAP values did not contain any features."}

    feature_names = ensure_feature_names_length(feature_names, shap_values.shape[1])
    friendly_names = [friendly_feature_label(name) for name in feature_names]
    base_feature_labels = [base_feature_label(name) for name in feature_names]

    mean_abs = np.abs(shap_values).mean(axis=0)
    aggregated = {}
    for idx, base in enumerate(base_feature_labels):
        aggregated[base] = aggregated.get(base, 0.0) + float(mean_abs[idx])

    aggregated_items = sorted(aggregated.items(), key=lambda item: item[1], reverse=True)
    global_rankings = [
        {"rank": rank, "feature": label, "mean_abs_shap": float(value)}
        for rank, (label, value) in enumerate(aggregated_items, 1)
    ]
    top_features = global_rankings[:SHAP_MAX_DISPLAY]

    SHAP_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    slug = slugify(model_name)
    plot_paths = []

    warning_message = "The NumPy global RNG was seeded by calling `np.random.seed`"

    try:
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message=warning_message, category=FutureWarning)
            shap.summary_plot(
                shap_values,
                data_for_plot,
                feature_names=friendly_names,
                max_display=SHAP_MAX_DISPLAY,
                show=False,
            )
        beeswarm_path = SHAP_OUTPUT_DIR / f"{slug}_shap_beeswarm.png"
        plt.gcf().savefig(beeswarm_path, bbox_inches="tight")
        plt.close()
        plot_paths.append(str(beeswarm_path))
    except Exception as exc:
        plt.close()
        plot_paths.append(f"Failed to generate SHAP beeswarm plot: {exc}")

    if top_features:
        try:
            bar_path = SHAP_OUTPUT_DIR / f"{slug}_shap_bar.png"
            fig, ax = plt.subplots(figsize=(8, max(3, len(top_features) * 0.4)))
            labels = [item["feature"] for item in top_features]
            values = [item["mean_abs_shap"] for item in top_features]
            positions = np.arange(len(labels))
            ax.barh(positions, values, color="#1f77b4")
            ax.set_yticks(positions)
            ax.set_yticklabels(labels)
            ax.invert_yaxis()
            ax.set_xlabel("mean |SHAP|")
            ax.set_ylabel("Feature")
            ax.set_title(f"{model_name} - SHAP importance")
            fig.tight_layout()
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", message=warning_message, category=FutureWarning)
                fig.savefig(bar_path, bbox_inches="tight")
            plt.close(fig)
            plot_paths.append(str(bar_path))
        except Exception as exc:
            if 'fig' in locals():
                plt.close(fig)
            else:
                plt.close()
            plot_paths.append(f"Failed to generate SHAP bar plot: {exc}")
    else:
        plot_paths.append("No SHAP features to visualise.")

    return {
        "name": model_name,
        "sample_size": int(X_sample.shape[0]),
        "top_features": top_features,
        "global_rankings": global_rankings,
        "plot_paths": plot_paths,
    }



scoring = {
    "mae": "neg_mean_absolute_error",
    "mse": "neg_mean_squared_error",
    "median_ae": "neg_median_absolute_error",
    "r2": "r2",
}

cv_search = KFold(n_splits=5, shuffle=True, random_state=42)
cv_eval = KFold(n_splits=10, shuffle=True, random_state=42)

model_configs = [
    {
        "name": "SVR (RBF)",
        "estimator": build_pipeline(SVR(kernel="rbf")),
        "search": "grid",
        "params": {
            "regressor__C": [1.0, 10.0, 100.0],
            "regressor__epsilon": [0.01, 0.1, 1.0],
            "regressor__gamma": ["scale", "auto"],
        },
        "n_jobs": -1,
        "cv_n_jobs": -1,
    },
    {
        "name": "Linear Regression",
        "estimator": build_pipeline(LinearRegression()),
        "search": None,
        "cv_n_jobs": -1,
    },
    {
        "name": "Random Forest",
        "estimator": build_pipeline(
            RandomForestRegressor(random_state=42, n_jobs=-1)
        ),
        "search": "random",
        "params": {
            "regressor__n_estimators": [100, 200, 400],
            "regressor__max_depth": [None, 10, 20, 30],
            "regressor__min_samples_split": [2, 5, 10],
            "regressor__min_samples_leaf": [1, 2, 4],
            "regressor__max_features": ["sqrt", "log2", None],
        },
        "n_iter": 12,
        "n_jobs": -1,
        "cv_n_jobs": -1,
    },
    {
        "name": "MLP",
        "estimator": build_pipeline(
            MLPRegressor(
                hidden_layer_sizes=(64, 32),
                activation="relu",
                solver="adam",
                learning_rate="adaptive",
                learning_rate_init=0.001,
                max_iter=200,
                random_state=42,
                early_stopping=True,
                n_iter_no_change=15,
            ),
            scale=True,
        ),
        "search": "grid",
        "params": {
            "regressor__hidden_layer_sizes": [(64, 32), (128, 64)],
            "regressor__alpha": [0.0001, 0.001, 0.01],
            "regressor__learning_rate_init": [0.001, 0.005],
        },
        "n_jobs": 1,
        "cv_n_jobs": 1,
    },
]

if LIGHTGBM_AVAILABLE:
    model_configs.append(
        {
            "name": "LightGBM",
            "estimator": build_pipeline(
                LGBMRegressor(
                    random_state=42,
                    n_estimators=300,
                    learning_rate=0.05,
                    subsample=0.9,
                    colsample_bytree=0.9,
                    n_jobs=-1,
                )
            ),
            "search": None,
            "cv_n_jobs": -1,
        }
    )

if TABNET_AVAILABLE:
    model_configs.append(
        {
            "name": "TabNet",
            "estimator": build_pipeline(
                TabNetRegressorWrapper(
                    n_d=16,
                    n_a=16,
                    n_steps=4,
                    gamma=1.5,
                    lambda_sparse=1e-3,
                    seed=42,
                    verbose=0,
                ),
                to_float32=True,
            ),
            "search": "random",
            "params": {
                "regressor__n_d": [8, 16, 24],
                "regressor__n_steps": [3, 4, 5],
                "regressor__gamma": [1.0, 1.5, 2.0],
                "regressor__lambda_sparse": [1e-4, 1e-3],
            },
            "n_iter": 4,
            "n_jobs": 1,
            "cv_n_jobs": 1,
        }
    )

SHAP_SAMPLE_SIZE = 500
SHAP_BACKGROUND_SIZE = 200
SHAP_MAX_DISPLAY = 15
SHAP_RANDOM_STATE = 42
SHAP_OUTPUT_DIR = Path("xai_outputs")

GLOBAL_IMPORTANCE_TOP_K = 10

search_summaries = []
cv_summaries = []
shap_summaries = []

for config in model_configs:
    name = config["name"]
    estimator = config["estimator"]
    search_type = config.get("search")
    best_estimator = estimator
    best_params = {}
    best_mae = None
    searcher = None

    if search_type == "grid":
        searcher = GridSearchCV(
            estimator,
            config["params"],
            cv=cv_search,
            scoring="neg_mean_absolute_error",
            n_jobs=config.get("n_jobs", 1),
        )
    elif search_type == "random":
        searcher = RandomizedSearchCV(
            estimator,
            config["params"],
            n_iter=config.get("n_iter", 10),
            cv=cv_search,
            scoring="neg_mean_absolute_error",
            n_jobs=config.get("n_jobs", 1),
            random_state=42,
        )

    if searcher is not None:
        searcher.fit(X, y)
        best_estimator = searcher.best_estimator_
        best_params = searcher.best_params_
        best_mae = -searcher.best_score_

    search_summaries.append(
        {
            "name": name,
            "tuned": searcher is not None,
            "best_params": best_params,
            "best_mae": best_mae,
            "search_type": search_type,
            "search_params": config.get("params") if searcher is not None else None,
            "n_iter": config.get("n_iter") if search_type == "random" else None,
        }
    )

    cv_scores = cross_validate(
        best_estimator,
        X,
        y,
        cv=cv_eval,
        scoring=scoring,
        n_jobs=config.get("cv_n_jobs", 1),
        error_score="raise",
    )

    cv_summaries.append(
        {
            "name": name,
            "mae": -cv_scores["test_mae"],
            "rmse": np.sqrt(-cv_scores["test_mse"]),
            "median_ae": -cv_scores["test_median_ae"],
            "r2": cv_scores["test_r2"],
        }
    )

    best_estimator.fit(X, y)
    raw_feature_names = get_pipeline_feature_names(best_estimator)

    if SHAP_AVAILABLE:
        shap_summary = generate_shap_summary(best_estimator, raw_feature_names, X, name)
        if shap_summary is not None:
            shap_summaries.append(shap_summary)

print("Hyperparameter search (MAE minimization)")
for summary in search_summaries:
    name = summary["name"]
    if summary["tuned"]:
        search_type = summary.get("search_type")
        if search_type == "grid":
            header = f"{name} | Grid search space:"
        elif search_type == "random":
            iter_info = f" (n_iter={summary.get('n_iter')})" if summary.get("n_iter") else ""
            header = f"{name} | Randomized search space{iter_info}:"
        else:
            header = f"{name} | Hyperparameter search space:"
        print(header)
        params = summary.get("search_params") or {}
        for line in describe_search_space(params):
            print(line)
        print(f"  Best CV MAE: {summary['best_mae']:,.2f}")
        for key in sorted(summary["best_params"], key=lambda k: k):
            print(f"    best {key}: {summary['best_params'][key]}")
    else:
        print(f"{name} | no hyperparameter search (defaults used)")
    print()

print("10-fold cross-validation with tuned models (mean +/- std)")
for metrics in cv_summaries:
    mae_scores = metrics["mae"]
    rmse_scores = metrics["rmse"]
    median_ae_scores = metrics["median_ae"]
    r2_scores = metrics["r2"]

    print(metrics["name"])
    print(f"  MAE mean: {mae_scores.mean():,.2f} | std: {mae_scores.std(ddof=1):,.2f}")
    print(f"  RMSE mean: {rmse_scores.mean():,.2f} | std: {rmse_scores.std(ddof=1):,.2f}")
    print(f"  Median AE mean: {median_ae_scores.mean():,.2f} | std: {median_ae_scores.std(ddof=1):,.2f}")
    print(f"  R^2 mean: {r2_scores.mean():.4f} | std: {r2_scores.std(ddof=1):.4f}")
    print()



if SHAP_AVAILABLE:
    if shap_summaries:
        print("SHAP (permutation) global feature importance (mean |SHAP| values)")
        for summary in shap_summaries:
            if "error" in summary:
                print(f"{summary['name']} | SHAP error: {summary['error']}")
                print()
                continue
            print(summary["name"])
            print(f"  Samples used: {summary['sample_size']}")
            top_items = summary.get("top_features", [])[:GLOBAL_IMPORTANCE_TOP_K]
            for item in top_items:
                print(f"  {item['rank']:>2}. {item['feature']} | mean |SHAP|: {item['mean_abs_shap']:.6f}")
            for path in summary["plot_paths"]:
                print(f"    Output: {path}")
            print()
    else:
        print("SHAP available but no SHAP summaries were generated.")
else:
    print("SHAP not available: install shap and matplotlib to generate SHAP plots.")

if not LIGHTGBM_AVAILABLE:
    print("LightGBM not available: install lightgbm to include it in the comparison.")

if not TABNET_AVAILABLE:
    print("TabNet not available: install pytorch-tabnet to include it in the comparison.")
