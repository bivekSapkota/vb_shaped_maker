# vb_shaped_maker
Creates a shaped file for gas

Makes a shaped file for gas using an existing datas, also known as dcs spreadsheet file.
utilizes reference id and prefix to search the deal and find volumes for that particular date

creates a new spreadsheet with the name of a deal and fills the hardcoded columns  with the volume for the particular deal
saves the file as csv with the file named as the reference id

limitations for now:
  1. only search the reference ids if the header in dcs file is starting from first row of first column
  2. is valid only if the volume obtained is for the whole month
  3. catching errors not implemented as of now
  4. works only for gas not for power
  5. not every variable is declared as of yet explicitly
