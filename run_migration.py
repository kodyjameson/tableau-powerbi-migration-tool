import subprocess

while True:
    print("\nTableau to Power BI Migration Tool")
    print("1 - Parse Tableau workbook")
    print("2 - Enhance migration workbook")
    print("3 - Exit")

    choice = input("Choose an option: ").strip()

    if choice == "1":
        subprocess.run(["python3", "tableau_to_powerbi_migration.py"])
    elif choice == "2":
        subprocess.run(["python3", "enhance_migration_workbook.py"])
    elif choice == "3":
        print("Goodbye.")
        break
    else:
        print("Invalid choice. Try again.")
