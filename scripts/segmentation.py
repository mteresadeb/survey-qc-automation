import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import silhouette_score
from pathlib import Path

# ─────────────────────────────────────────
# CONFIGURATION
# Adapt these variables to your project
# ─────────────────────────────────────────

INPUT_FILE  = Path('./data/survey_data.csv')
OUTPUT_DIR  = Path('./output')
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Attitudinal/behavioral columns to use for clustering
PSYCHOMETRIC_VARS = [
    'sustainability_concern',
    'price_sensitivity',
    'brand_loyalty',
    'status_orientation',
    'digital_engagement',
    'ethical_consumption',
    'convenience_priority',
    'trend_sensitivity',
]

K            = 4
RANDOM_STATE = 42

SEGMENT_COLORS = {
    'Conscious Consumer':    '#2E86AB',
    'Practical Consumer':    '#A23B72',
    'Aspirational Consumer': '#F18F01',
    'Traditional Consumer':  '#4CAF50',
}


# ─────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────

def load_data(path):
    df = pd.read_csv(path, low_memory=False)
    missing = [v for v in PSYCHOMETRIC_VARS if v not in df.columns]
    if missing:
        raise ValueError(f'Missing columns: {missing}')
    print(f'Data loaded: {df.shape[0]} rows, {df.shape[1]} columns')
    return df


# ─────────────────────────────────────────
# SCALING
# ─────────────────────────────────────────

def scale_features(df):
    scaler = StandardScaler()
    X = scaler.fit_transform(df[PSYCHOMETRIC_VARS])
    print(f'Features scaled: {X.shape}')
    return X, scaler


# ─────────────────────────────────────────
# OPTIMAL K SELECTION
# ─────────────────────────────────────────

def find_optimal_k(X, k_min=2, k_max=10):
    inertias, silhouettes = [], []
    k_range = range(k_min, k_max + 1)

    for k in k_range:
        km = KMeans(n_clusters=k, random_state=RANDOM_STATE, n_init=10)
        labels = km.fit_predict(X)
        inertias.append(km.inertia_)
        silhouettes.append(silhouette_score(X, labels))

    fig, axes = plt.subplots(1, 2, figsize=(12, 4))

    axes[0].plot(list(k_range), inertias, marker='o', color='#378ADD', linewidth=2)
    axes[0].axvline(K, color='#E24B4A', linestyle='--', linewidth=1.5, label=f'Selected k={K}')
    axes[0].set_xlabel('k')
    axes[0].set_ylabel('Inertia')
    axes[0].set_title('Elbow Method')
    axes[0].legend()

    axes[1].plot(list(k_range), silhouettes, marker='o', color='#4CAF50', linewidth=2)
    axes[1].axvline(K, color='#E24B4A', linestyle='--', linewidth=1.5, label=f'Selected k={K}')
    axes[1].set_xlabel('k')
    axes[1].set_ylabel('Silhouette Score')
    axes[1].set_title('Silhouette Score')
    axes[1].legend()

    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'optimal_k.png', dpi=150, bbox_inches='tight')
    plt.show()

    print(f'\nSilhouette at k={K}: {silhouettes[K - k_min]:.4f}')
    return k_range, inertias, silhouettes


# ─────────────────────────────────────────
# CLUSTERING
# ─────────────────────────────────────────

def fit_segments(X, scaler):
    km = KMeans(n_clusters=K, random_state=RANDOM_STATE, n_init=10)
    clusters = km.fit_predict(X)

    centroids_df = pd.DataFrame(
        scaler.inverse_transform(km.cluster_centers_),
        columns=PSYCHOMETRIC_VARS
    ).round(2)

    def assign_label(row):
        scores = {
            'Conscious Consumer':    row['sustainability_concern'] + row['ethical_consumption'],
            'Practical Consumer':    row['price_sensitivity'] + row['convenience_priority'],
            'Aspirational Consumer': row['status_orientation'] + row['trend_sensitivity'],
            'Traditional Consumer':  row['brand_loyalty'] + (5 - row['digital_engagement']),
        }
        return max(scores, key=scores.get)

    centroids_df['segment'] = centroids_df.apply(assign_label, axis=1)
    cluster_to_label = centroids_df['segment'].to_dict()

    print('Cluster to segment mapping:')
    for k, v in cluster_to_label.items():
        print(f'  Cluster {k} -> {v}')

    return clusters, cluster_to_label, centroids_df


# ─────────────────────────────────────────
# PROFILING
# ─────────────────────────────────────────

def profile_segments(df):
    profile = df.groupby('segment')[PSYCHOMETRIC_VARS].mean().round(2)
    print('\nSegment profiles (mean scores):')
    print(profile)
    return profile


# ─────────────────────────────────────────
# VISUALIZATIONS
# ─────────────────────────────────────────

def plot_radar(profile):
    labels_radar = [v.replace('_', '\n') for v in PSYCHOMETRIC_VARS]
    num_vars = len(PSYCHOMETRIC_VARS)
    angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    for segment in profile.index:
        values = profile.loc[segment].tolist()
        values += values[:1]
        color = SEGMENT_COLORS.get(segment, 'grey')
        ax.plot(angles, values, linewidth=2, color=color, label=segment)
        ax.fill(angles, values, alpha=0.08, color=color)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels_radar, fontsize=8)
    ax.set_ylim(1, 5)
    ax.set_title('Segment profiles', fontsize=12, pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.35, 1.1), fontsize=9)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'radar_chart.png', dpi=150, bbox_inches='tight')
    plt.show()

def plot_pca_map(df, X):
    pca = PCA(n_components=2, random_state=RANDOM_STATE)
    X_pca = pca.fit_transform(X)
    df = df.copy()
    df['pca1'] = X_pca[:, 0]
    df['pca2'] = X_pca[:, 1]

    fig, ax = plt.subplots(figsize=(9, 6))
    for segment in df['segment'].unique():
        mask = df['segment'] == segment
        color = SEGMENT_COLORS.get(segment, 'grey')
        ax.scatter(df.loc[mask, 'pca1'], df.loc[mask, 'pca2'],
                   label=segment, color=color, alpha=0.45, s=18)

    ax.set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.1f}% variance)')
    ax.set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.1f}% variance)')
    ax.set_title('PCA map — consumer segments', fontsize=12)
    ax.legend(fontsize=9)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'pca_map.png', dpi=150, bbox_inches='tight')
    plt.show()

def plot_size(df):
    size = df['segment'].value_counts().reindex(list(SEGMENT_COLORS.keys()))
    pcts = (size / size.sum() * 100).round(1)
    colors_bar = [SEGMENT_COLORS[s] for s in size.index]

    fig, ax = plt.subplots(figsize=(8, 4))
    bars = ax.barh(size.index, size.values, color=colors_bar, edgecolor='white', height=0.6)
    for bar, pct in zip(bars, pcts):
        ax.text(bar.get_width() + 5, bar.get_y() + bar.get_height() / 2,
                f'{pct}%', va='center', fontsize=10)
    ax.set_xlabel('Number of respondents')
    ax.set_title('Segment size distribution', fontsize=12)
    ax.invert_yaxis()
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / 'segment_size.png', dpi=150, bbox_inches='tight')
    plt.show()


# ─────────────────────────────────────────
# EXPORT
# ─────────────────────────────────────────

def export_results(df, profile):
    export_cols = ['segment'] + PSYCHOMETRIC_VARS
    df[export_cols].to_csv(OUTPUT_DIR / 'segmented_data.csv', index=False)
    profile.to_csv(OUTPUT_DIR / 'segment_profiles.csv')
    print(f'\nResults saved to: {OUTPUT_DIR}')
    print('  segmented_data.csv   -- full dataset with segment assignments')
    print('  segment_profiles.csv -- mean scores per segment')


# ─────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────

def run_segmentation(input_file=INPUT_FILE):
    print('=== Audience Segmentation Pipeline ===\n')

    df                               = load_data(input_file)
    X, scaler                        = scale_features(df)
    find_optimal_k(X)
    clusters, cluster_to_label, _    = fit_segments(X, scaler)

    df['cluster'] = clusters
    df['segment'] = df['cluster'].map(cluster_to_label)

    profile = profile_segments(df)
    plot_radar(profile)
    plot_pca_map(df, X)
    plot_size(df)
    export_results(df, profile)

    print('\n=== Pipeline complete ===')
    return df, profile


if __name__ == '__main__':
    run_segmentation()
