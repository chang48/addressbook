# Simple LaTeX Address Book

A simple workflow that generates a LaTeX address book. The workflow consists of the following elements
- Google Forms
  - Each entry of the address book represents a household. So the form has a common section "Household Address" that collects household
    address info such as: street number, city, state, and zip code.
  - Following the "Household Address" section are multiple member contact cards. Each card collects name, phone number, and email
    address.
- Google Sheet
  - Collects Google Forms responses
  - Use App Script to convert multi-section responses into rows with the common "Household Address" column. Because the Household Address
    is unique, this column will be used in pandas `groupby` operation, as explained in the following.
- A python script that pulls the rows and write the contact info into `.tex` file for each addressbook entry. Here the converted responses
  are imported as a pandas dataframe. Using the Household Address as `groupby` column, the code iterates over each household and outputs
  address book entry in `.tex` format.
 
