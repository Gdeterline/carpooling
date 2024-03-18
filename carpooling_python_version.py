# %%
# Import the required libraries
import pandas as pd


# Load the workbooks and their respective worksheets
offers_worksheet = pd.read_excel("./Offres_2022.xlsx", sheet_name="Matching")
requests_worksheet = pd.read_excel("./Demandes_2022.xlsx", sheet_name="Demandes")

print("Workbooks loaded successfully")

#########################################################################################
# %%


# Print the first 5 rows of the worksheet
offers_worksheet.head(5)                  # Seems to have issues. Need to check path and file name
requests_worksheet.head(5)                 # Seems to have issues. Need to check path and file name

#########################################################################################
# %%

# Check the columns of the worksheet
requests_worksheet.columns
offers_worksheet.columns


#########################################################################################
# %%

# First treatment of the offer sheet data

horaires_offers = offers_worksheet.columns[6:]
print(horaires_offers)
print()
name_offers = offers_worksheet.loc[:, "Nom"]
print(name_offers)
print()
child_offers = offers_worksheet.loc[:, "Enfant 1":"Enfant 3"]
print(child_offers)
print()
number_places_offers = offers_worksheet.loc[:, "Places"]
print(number_places_offers)
print()
availabilities_offers = offers_worksheet.loc[:, "Lundi 8h":"Vendredi 17h30"]
print(availabilities_offers)
print()


# Try push to a datasheet
# Convert horaires_offers to a DataFrame
#horaires_offers_df = pd.DataFrame(horaires_offers)
#horaires_offers_df.to_csv("./horaires_offers.csv", index=False, header=False)

# %%

# First treatment of the request sheet data

horaires_requests = requests_worksheet.columns[1:]
print(horaires_requests)
print()
name_requests = requests_worksheet.loc[:, "Prénom"]
print(name_requests)
print()
availabilities_requests = requests_worksheet.loc[:, "Lundi 8h":"Vendredi 18h30"]
print(availabilities_requests)
print()

# %%

# Get number of drivers and passengers

nb_requests = 0

number_drivers = len(name_offers)
number_passengers = len(name_requests)
print(f"Number of drivers: {number_drivers}")
print(f"Number of passengers: {number_passengers}")

# Print total number of requests over 2 weeks
for i in range(number_passengers):
    for j in range(availabilities_requests.shape[1]):
        if availabilities_requests.iloc[i, j] == "Oui":
            nb_requests += 2
        elif availabilities_requests.iloc[i, j] == "Pair":
            nb_requests += 1
        elif availabilities_requests.iloc[i, j] == "Impair":
            nb_requests += 1
            
        
print(f"Number of requests: {nb_requests}")
            

# %%


