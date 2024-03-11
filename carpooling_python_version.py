# %%
# Import the required libraries
import pandas as pd


# Load the workbooks and their respective worksheets
offers_worksheet = pd.read_excel("./Offres_2022.xlsx", sheet_name="Matching")
demands_worksheet = pd.read_excel("./Demandes_2022.xlsx", sheet_name="Demandes")

print("Workbooks loaded successfully")

#########################################################################################
# %%


# Print the first 5 rows of the worksheet
offers_worksheet.head(5)                  # Seems to have issues. Need to check path and file name
demands_worksheet.head(5)                 # Seems to have issues. Need to check path and file name

#########################################################################################
# %%

# Check the columns of the worksheet
demands_worksheet.columns
offers_worksheet.columns


#########################################################################################
# %%

horaires_offers = offers_worksheet.columns[6:]
print(horaires_offers)
name_offers = offers_worksheet.loc[:, "Nom"]
print(name_offers)
child_offers = offers_worksheet.loc[:, "Enfant 1":"Enfant 3"]
print(child_offers)
number_places_offers = offers_worksheet.loc[:, "Places"]
print(number_places_offers)
availabilities_offers = offers_worksheet.loc[:, "Lundi 8h":"Vendredi 17h30"]
print(availabilities_offers)


# %%
