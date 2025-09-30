import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt
#pip install openpyxl
from rapidfuzz import process, fuzz
import os

# Set working directory
os.chdir(r"C:\Users\jadhav\OneDrive - The University of Chicago\Desktop\Busara\democracy\input")

# Load Excel file
df = pd.read_excel("lit_review_data.xlsx", sheet_name="Graphics_data")

# Clean whitespace from 'Country' column
df['Country'] = df['Country'].str.strip()

# Count frequency
country_counts = df['Country'].value_counts().reset_index()
country_counts.columns = ['Country', 'Frequency']

# Load world shapefile
world = gpd.read_file("110m_cultural/ne_110m_admin_0_countries.shp")
official_names = world['ADMIN'].unique()

# Fuzzy match function
def fuzzy_match_country(name, choices, scorer=fuzz.WRatio, threshold=70):
    match = process.extractOne(name, choices, scorer=scorer)
    if match and match[1] >= threshold:
        return match[0]
    else:
        return None

# Apply fuzzy matching
country_counts['Matched_Country'] = country_counts['Country'].apply(lambda x: fuzzy_match_country(x, official_names))

# Print or export unmatched values
unmatched = country_counts[country_counts['Matched_Country'].isnull()]
print("Unmatched countries:\n", unmatched)

# Drop unmatched for now
matched_counts = country_counts.dropna(subset=['Matched_Country'])

# Merge with world map
world = world.merge(matched_counts, how='left', left_on='ADMIN', right_on='Matched_Country')

import matplotlib.pyplot as plt

# Remove Antarctica
world_no_antarctica = world[world['ADMIN'] != 'Antarctica']

import matplotlib as mpl

# Plot 
fig, ax = plt.subplots(1, 1, figsize=(15, 10))
country_plot = world_no_antarctica.plot(
    column='Frequency',
    ax=ax,
    legend=False,  # Disable GeoPandas auto-legend
    cmap='Blues',  
    missing_kwds={"color": "lightgrey", "label": "No Data"},
    edgecolor='Black',
    linewidth=0.5
)

# Add a square outline around the map
xmin, xmax = ax.get_xlim()
ymin, ymax = ax.get_ylim()
rect = plt.Rectangle((xmin, ymin), xmax - xmin, ymax - ymin,
                     fill=False, color='black', linewidth=2)
ax.add_patch(rect)

# Add horizontal colorbar (legend)
norm = mpl.colors.Normalize(vmin=world_no_antarctica['Frequency'].min(),
                            vmax=world_no_antarctica['Frequency'].max())
cbar = fig.colorbar(
    mpl.cm.ScalarMappable(norm=norm, cmap='Blues'),
    ax=ax,
    orientation='horizontal',
    shrink=0.8,
    pad=0.05
)
cbar.set_label('Frequency')

ax.set_title('Heatmap of Countries Studied in Literature Review', fontsize=18, fontweight='bold')
ax.axis('off')
plt.tight_layout()
plt.show()

################################################################################
# Map 2 by Continent  
################################################################################

# Clean whitespace and casing from 'Continent' column
df['Continent'] = df['Continent'].str.strip().str.title()

# Count frequency of each continent
continent_counts = df['Continent'].value_counts().reset_index()
continent_counts.columns = ['CONTINENT', 'FREQUENCY']

# Move Russia from Europe to Asia BEFORE dissolving
world.loc[(world['ADMIN'] == 'Russia'), 'CONTINENT'] = 'Asia'

# Dissolve countries into continents
continents = world.dissolve(by='CONTINENT', as_index=False)

# Merge frequencies with the continent geometries
continents = continents.merge(continent_counts, on='CONTINENT', how='left')
continents['FREQUENCY'] = continents['FREQUENCY'].fillna(0)

continents_no_antarctica = continents[continents['CONTINENT'] != 'Antarctica']

# Plot
fig, ax = plt.subplots(1, 1, figsize=(15, 10))
continent_plot = continents_no_antarctica.plot(
    column='FREQUENCY',
    ax=ax,
    cmap='Blues',
    edgecolor='Black',
    linewidth=0.5,
    legend=False  # turn off GeoPandas default legend
)

# Add a square outline around the map
xmin, xmax = ax.get_xlim()
ymin, ymax = ax.get_ylim()
rect = plt.Rectangle((xmin, ymin), xmax - xmin, ymax - ymin,
                     fill=False, color='black', linewidth=2)
ax.add_patch(rect)

# Create custom colorbar (legend)
norm = mpl.colors.Normalize(vmin=continents_no_antarctica['FREQUENCY'].min(),
                            vmax=continents_no_antarctica['FREQUENCY'].max())
cbar = fig.colorbar(
    mpl.cm.ScalarMappable(norm=norm, cmap='Blues'),
    ax=ax,
    orientation='horizontal',    # makes it horizontal
    shrink=0.8,                  # smaller size (0.6 = 60%)
    pad=0.05                     # space between map and legend
)

cbar.set_label('Frequency')

# Final touches
ax.set_title('Heatmap of Continents Studied in Literature Review', fontsize=18, fontweight='bold')
ax.axis('off')
plt.tight_layout()
plt.show()

################################################################################
# Tables   
################################################################################

# Load Excel file
df_theme = pd.read_excel("lit_review_data.xlsx", sheet_name="Themes")

# Drop rows with missing regions if desired
df_clean = df_theme.dropna(subset=["Region"])

# Pivot table for Themes
theme_pivot = (
    df_clean
    .groupby(['Suggested Theme', 'Region'])
    .size()
    .unstack(fill_value=0)
    .reset_index()
    .rename_axis(None, axis=1)
)

# Optional: sort columns (first column is 'Suggested Theme', rest are continents)
cols = ['Suggested Theme'] + sorted([col for col in theme_pivot.columns if col != 'Suggested Theme'])
theme_pivot = theme_pivot[cols]

# Export to Excel
with pd.ExcelWriter(r"C:\Users\jadhav\OneDrive - The University of Chicago\Desktop\Busara\democracy\output\continent_theme_method_tables.xlsx", engine="openpyxl", mode="a" if os.path.exists(r"C:\Users\jadhav\OneDrive - The University of Chicago\Desktop\Busara\democracy\output\continent_theme_method_tables.xlsx") else "w") as writer:
    theme_pivot.to_excel(writer, sheet_name="Themes_vis", index=False)

# Print table
print("\nThemes by Continent:\n", theme_pivot)

# Clustered bar chart: Themes by Continent
import matplotlib.pyplot as plt

# Set index to 'Suggested Theme' for plotting
theme_pivot_plot = theme_pivot.set_index('Suggested Theme')

# Plot
ax = theme_pivot_plot.plot(kind='bar', figsize=(14, 7))
plt.title('Themes by Continent')
plt.ylabel('Count')
plt.xlabel('Suggested Theme')
plt.xticks(rotation=45, ha='right')
plt.legend(title='Continent')
plt.tight_layout()
plt.show()
  
# Methodology table 
# Pivot table for Themes
method_pivot = (
    df_clean
    .groupby(['Methodology', 'Region'])
    .size()
    .unstack(fill_value=0)
    .reset_index()
    .rename_axis(None, axis=1)
)

# Optional: sort columns (first column is 'Methodology', rest are continents)
cols = ['Methodology'] + sorted([col for col in method_pivot.columns if col != 'Methodology'])
method_pivot = method_pivot[cols]

# Export to Excel
with pd.ExcelWriter(r"C:\Users\jadhav\OneDrive - The University of Chicago\Desktop\Busara\democracy\output\continent_theme_method_tables.xlsx", engine="openpyxl", mode="a" if os.path.exists(r"C:\Users\jadhav\OneDrive - The University of Chicago\Desktop\Busara\democracy\output\continent_theme_method_tables.xlsx") else "w") as writer:
    method_pivot.to_excel(writer, sheet_name="Method_vis", index=False)

# Print table
print("\nThemes by Continent:\n", method_pivot)

################################################################################
# Map 3 Most frequent theme by continent   
################################################################################

# Sort columns
cols = ['Suggested Theme'] + sorted([col for col in theme_pivot.columns if col != 'Suggested Theme'])
theme_pivot = theme_pivot[cols]

# Reshape to long format
theme_long = theme_pivot.melt(id_vars='Suggested Theme', var_name='Continent', value_name='Frequency')
theme_long = theme_long[theme_long['Frequency'] > 0]

# Get most frequent theme per continent
top_themes = (
    theme_long.sort_values(['Continent', 'Frequency'], ascending=[True, False])
    .drop_duplicates(subset=['Continent'])
    .rename(columns={'Suggested Theme': 'Top_Theme'})
)

# Move Russia from Europe to Asia
world.loc[world['ADMIN'] == 'Russia', 'CONTINENT'] = 'Asia'

# Dissolve countries into continents
continents = world.dissolve(by='CONTINENT', as_index=False)

# Merge theme frequency data
continents = continents.merge(top_themes, left_on='CONTINENT', right_on='Continent', how='left')
continents['Frequency'] = continents['Frequency'].fillna(0)

# Remove Antarctica
continents_no_antarctica = continents[continents['CONTINENT'] != 'Antarctica']

fig, ax = plt.subplots(1, 1, figsize=(16, 10))

# Plot base map
continents_no_antarctica.plot(ax=ax, color='lightgrey', edgecolor='black')

# Annotate with bubble and theme label
for idx, row in continents_no_antarctica.iterrows():
    if pd.notnull(row['Top_Theme']):
        point = row['geometry'].representative_point()
        # Theme label
        # Split the theme into two roughly equal parts for vertical stacking
        theme = row['Top_Theme']
        words = theme.split()
        if len(words) > 1:
            mid = len(words) // 2
            # If odd, put more words on the first line
            line1 = ' '.join(words[:mid + (len(words) % 2)])
            line2 = ' '.join(words[mid + (len(words) % 2):])
            label = f"{line1}\n{line2}"
        else:
            label = theme
        ax.text(point.x, point.y, label, fontsize=10, ha='center', va='center', color='black',
            bbox=dict(boxstyle='round,pad=0.2', fc='white', alpha=0.6, ec='none'), zorder=6)

# Aesthetics
plt.title('Most Explored Theme by Continent', fontsize=16)
ax.axis('off')
plt.tight_layout()
plt.show()

