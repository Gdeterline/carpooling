# %%
# Import the required libraries
import pandas as pd
import numpy as np


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
#print(horaires_offers)
print()
name_offers = offers_worksheet.loc[:, "Nom"]
print(name_offers)
print()
child_offers = offers_worksheet.loc[:, "Enfant 1":"Enfant 3"]
#print(child_offers)
print()
number_places_offers = offers_worksheet.loc[:, "Places"]
#print(number_places_offers)
print()
availabilities_offers = offers_worksheet.loc[:, "Lundi 8h":"Vendredi 17h30"]
#print(availabilities_offers)
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

# number of schedules T
T = requests_worksheet.shape[1] - 1
print("Nombre d'horaire T =", T)



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

# family matrix with the first column being the passenger number (sheet Demandes_2022) and the second column being the driver number (from the same family) (sheet Offres_2022)

family = np.zeros((number_passengers, 2), dtype=int)

for i in range(number_passengers):
    family[i, 0] = i + 1
    for j in range(number_drivers):
        if name_requests[i] in child_offers.iloc[j].values:
            family[i, 1] = j + 1

print(family)
print()


# %%

availabilities_offers.head(5)


# %%

# places_offered_drivers matrix with the first column being the driver number (sheet Offres_2022) and the second column being the number of places offered by the driver
# over the two weeks

places_offered_drivers = np.zeros((number_drivers, 2), dtype=int)

for i in range(number_drivers):
    places_offered_drivers[i, 0] = i + 1
    for j in range(availabilities_offers.shape[1]):
        if availabilities_offers.iloc[i, j] == "Oui": 
            places_offered_drivers[i, 1] += 2 * number_places_offers[i]
        elif availabilities_offers.iloc[i, j] == "Impair":
            places_offered_drivers[i, 1] += number_places_offers[i]
        elif availabilities_offers.iloc[i, j] == "Pair":
            places_offered_drivers[i, 1] += number_places_offers[i]

print(places_offered_drivers)
print()


#%%

# places_offered_passengers matrix with the first column being the passenger number (sheet Demandes_2022) and the second column being the number of places offered by the driver

places_offered_passengers = np.zeros((number_passengers, 2), dtype=int)

for i in range(number_passengers):
    places_offered_passengers[i, 0] = i + 1
    for j in range(number_drivers):
        if family[i, 1] == places_offered_drivers[j, 0]:
            places_offered_passengers[i, 1] = places_offered_drivers[j, 1]

print(places_offered_passengers)
print()


# %%


print(availabilities_offers.stack().unique())

# %%

# Determine the number of offers over one week (equals two weeks here since there are only Oui and Non values in the availabilities_offers matrix)
N_o = sum(x in ["Oui", "Pair", "Impair"] for x in availabilities_offers.values.flatten())
print(f"Number of offers: {N_o}")

# Determine the number of requests over one week
N_r = sum(x in ["Oui", "Pair", "Impair"] for x in availabilities_requests.values.flatten())
print(f"Number of requests: {N_r}")


# %%

# Create the matrices O and R

# Initialize matrix O
O = np.zeros((N_o, 5), dtype=int)

# Counter for valid offers
idx = 0


# Traverse the offers array

for i in range(availabilities_offers.shape[0]):
    nb_places = number_places_offers[i]
    for j in range(availabilities_offers.shape[1]):
        if availabilities_offers.iloc[i, j] == "Oui":
            # Store the offer in matrix O
            O[idx, 0] = i + 1
            O[idx, 1] = j + 1
            O[idx, 2] = 1
            O[idx, 3] = 1
            O[idx, 4] = nb_places
            idx += 1
        elif availabilities_offers.iloc[i, j] == "Pair":
            # Store the offer in matrix O
            O[idx, 0] = i + 1
            O[idx, 1] = j + 1
            O[idx, 2] = 1
            O[idx, 3] = 0
            O[idx, 4] = nb_places
            idx += 1
        elif availabilities_offers.iloc[i, j] == "Impair":
            # Store the offer in matrix O
            O[idx, 0] = i + 1
            O[idx, 1] = j + 1
            O[idx, 2] = 0
            O[idx, 3] = 1
            O[idx, 4] = nb_places
            idx += 1


print("Matrix O:")
print(O)
print()



# Initialize matrix R
R = np.zeros((N_r, 4), dtype=int)

# Counter for valid requests
idr = 0

# Traverse the requests array
for i in range(availabilities_requests.shape[0]):
    for j in range(availabilities_requests.shape[1]):
        if availabilities_requests.iloc[i, j] == "Oui":
            # Store the request in matrix R
            R[idr, 0] = i + 1
            R[idr, 1] = j + 1
            R[idr, 2] = 1
            R[idr, 3] = 1
            idr += 1
        elif availabilities_requests.iloc[i, j] == "Pair":
            # Store the request in matrix R
            R[idr, 0] = i + 1
            R[idr, 1] = j + 1
            R[idr, 2] = 1
            R[idr, 3] = 0
            idr += 1
        elif availabilities_requests.iloc[i, j] == "Impair":
            # Store the request in matrix R
            R[idr, 0] = i + 1
            R[idr, 1] = j + 1
            R[idr, 2] = 0
            R[idr, 3] = 1
            idr += 1

print("Matrix R:")
print(R)
print()

#%%
# Display drivers and passengers


# Traverse all days and hours
for t in range(1, T+1):
    # Find available drivers at the hour and even weeks
    available_drivers_pairs = np.unique(O[(O[:, 1] == t) & (O[:, 2] == 1), 0])
    # Find available drivers at the hour and odd weeks
    available_drivers_impairs = np.unique(O[(O[:, 1] == t) & (O[:, 3] == 1), 0])

    # Display available drivers at the hour and even weeks
    print(f"Drivers available for a ride at hour {t} on even weeks:")
    print(available_drivers_pairs)
    
    # Display available drivers at the hour and odd weeks
    print(f"Drivers available for a ride at hour {t} on odd weeks:")
    print(available_drivers_impairs)

print()

for t in range(1, T+1):
    # Find available passengers at the hour and even weeks
    available_passengers_pairs = np.unique(R[(R[:, 1] == t) & (R[:, 2] == 1), 0])
    # Find available passengers at the hour and odd weeks
    available_passengers_impairs = np.unique(R[(R[:, 1] == t) & (R[:, 3] == 1), 0])

    # Display available passengers at the hour and even weeks
    print(f"Passengers requesting a ride at hour {t} on even weeks:")
    print(available_passengers_pairs)
    
    # Display available passengers at the hour and odd weeks
    print(f"Passengers requesting a ride at hour {t} on odd weeks:")
    print(available_passengers_impairs)


# %%

# A and B initialization with zeros
A = np.zeros((number_drivers, T, 2), dtype=int)
B = np.zeros((number_passengers, T, 2), dtype=int)

###### Reminder : Julia's indexes begin at 1, Python's start at 0

# Fill A with O values
for i in range(O.shape[0]):
    driver = O[i, 0] - 1  # Adjust the index by subtracting 1
    time = O[i, 1] - 1  # Adjust the index by subtracting 1
    if O[i, 2] == 1:
        A[driver, time, 0] = 1
    if O[i, 3] == 1:
        A[driver, time, 1] = 1


# Fill B with R values
for i in range(R.shape[0]):
    passenger = R[i, 0] - 1  # Adjust the index by subtracting 1
    time = R[i, 1] - 1  # Adjust the index by subtracting 1
    if R[i, 2] == 1:
        B[passenger, time, 0] = 1
    if R[i, 3] == 1:
        B[passenger, time, 1] = 1


print("Variable A:")
print(A)

# For details on the A matrix uncomment the following
print(len(A))
print(len(A[0]))
with np.printoptions(threshold=np.inf):
    print(A)

#%%

# print(A[:,:,1]) To see the first part of the full A matrix

#%%


print("Variable A:")
print(B)

# For details on the B matrix uncomment the following
print(len(B))
print(len(B[0]))
with np.printoptions(threshold=np.inf):
    print(B)


# %%

# Create the M matrix
M = np.zeros((number_passengers, number_drivers, T, 2), dtype=float)

Beta = 7

for parent in range(number_drivers):
    for child in range(number_passengers):
        #Need condition for :  if name_requests[child] == child_offers[parent]
        if name_requests[child] in child_offers.iloc[parent].values: #### To check
            for t in range(T):
                for w in range(2):
                    M[child, parent, t, w] = Beta
        else:
            for t in range(T):
                for w in range(2):
                    M[child, parent, t, w] = 10 - Beta

print("Matrix M:")
with np.printoptions(threshold=np.inf):
    print(M)


# %%
