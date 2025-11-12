import bw2data as bd
import bw2io as bi
import bw2calc as bc
import pandas as pd
from bw2data import Database

### Import LCI from Excel into Brightway
########################################################

print(bd.projects)

# Define variables first
projectName = "paw_lca"
database = "paw_db" # check
ei_match = "ecoinvent-3.11-cutoff"  # the EI version in the project  & file
biosphere_match = "ecoinvent-3.11-biosphere"  # the biosphere version in the project  & file
path = "LCI_ Inventories"
# r"C:\Users\Asus\OneDrive - Universidade de Santiago de Compostela\07_Papers\Harmonization study\harmonization inventories_importer_db_test.xlsx"
##########################################################

# Ensure the Brightway project is set
if projectName in bd.projects:
    bd.projects.set_current(projectName)
    print(f"Project {projectName} is set as current project.")
else:
    print(f"Project {projectName} does not exist.")
    exit()
print("Available databases in the project:", list(bd.databases))
# check databases in the project
if database in bd.databases:
    print(f"Database {database} is already in the project. Will add new data to it.")
    user_input = input("Do you want to continue? (y/n): ").strip().lower()
    if user_input == "n":
        exit()
else:
    user_input = (
        input(f"Do you want to import the database {database}? (y/n): ").strip().lower()
    )
    if user_input == "y":
        print(f"Database {database} will be imported.")
    else:
        print("Process stopped by user.")
        exit()

########################################################
# Create the importer
harmonization_db = bi.ExcelImporter(path)

# apply strategies for matching
harmonization_db.apply_strategies()
# match database to itself
harmonization_db.match_database(fields=["name", "unit", "location"])
# match database to ecoinvent
harmonization_db.match_database(
    ei_match, fields=["name", "location", "reference product"]
)

# match database to biopshere flows
harmonization_db.match_database(biosphere_match, fields=["name", "categories"])
# harmonization_db.match_database("pharma-biosphere", fields=["name", "categories", "location"])

# Optionally, you can inspect unmatched exchanges or activities
print(harmonization_db.statistics())

if harmonization_db.all_linked:
    print("All flows are linked. Database will be written.")
    harmonization_db.write_database()
    print("New database created: ", list(bd.databases))

    # Check how many activities were imported
    db = bd.Database(database)
    activity_count = len(list(db))
    print(f"Number of activities imported: {activity_count}")

    # List all activities
    print("Activities in the database:")
    for i, act in enumerate(db, 1):
        print(f"{i}. {act['name']} ({act['location']})")

else:
    print("Some flows are not linked.")
    print("Unlinked flows:")
    for flow in harmonization_db.unlinked:
        print(flow)
    exit()


# Now your database is available in brightway
# check metadata
bd.Database(database).metadata

# ask if you want to keep database
user_input = input("Do you want to keep the database? (y/n): ").strip().lower()
if user_input == "n":
    del bd.databases[database]
    bd.databases
    print("Available databases in Project:", projectName)
    print(bd.databases)
    exit()
else:
    pass
