# How to Run the Carpooling Optimization Script

This guide will walk you through the steps to run the carpooling optimization script on your computer, whether you're using macOS, Linux, or Windows.

## Step 1: Install Python

Ensure you have Python installed on your computer. You can download it from python.org.

## Step 2: Install Required Libraries

The script requires specific Python libraries. Open your terminal (macOS/Linux) or Command Prompt (Windows) and run the following commands:

**pip install pandas numpy openpyxl pulp**


## Step 3: Prepare Your Excel Files

Make sure you have the following Excel files in the same directory as your script:

- Demandes_2022.xlsx (Requests)
- Offres_2022.xlsx (Offers)

(Note that 2022 is the year we worked on. Please adapt to the year you are in)

## Step 4: Adapt The Name Of The Input Files To Yours

In the "carpooling_python_version.py" file, you will have 2 things to change yourself.


On lines 18 and 19 of the file, you'll find the following codelines :

\# Load the workbooks and their respective worksheets
input_file = "*./Offres_2022.xlsx*"
offers_worksheet = pd.read_excel("./*Offres_2022.xlsx*", sheet_name="Réponses au formulaire 1")
requests_worksheet = pd.read_excel("./*Demandes_2022.xlsx*", sheet_name="Réponses au formulaire 1")

You will need to replace both *Offres_2022.xlsx* and *Demandes_2022.xlsx* by the name of your matching files. 
The name of your file must contain the year you are in, on the same model as the names above.

## Step 5 : Run the script

- Console Output: The console will display whether an optimal solution was found.
- Repartition_Voiture_2022_vf.xlsx: This new Excel file will be generated with the formatted data.



#### Drive safe !