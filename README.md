# spro_update_user

Program and shell scripts to automatically update JWSNY school pro user list at the beginning of a school year.

# Program and shell script files

* **WORK/update_user.py**    The python program to process inputs and generate new school pro user list. Inputs/outputs are basically in CSV format in UTF8 encoding.
* **WORK/update_user.sh**    WSL2 wrapper shell script to run the program, along with converting Japanese encoding.

* **update_user.bat**        Windows wrapper batch file, simply execute the above update_user.sh on WSL2.

# Input CSV files (not included in this git repo)

Please prepare for and copy the following input files under the InputCSV/ directory:

* **school_student_list.csv**    The list of the students of current school year, provided by the school
* **spro_current_user.csv**   The current snapshot of the existing users, exported from School Pro.

# How to run the program

After you place the above input CSV files, please execute (double click) the udpate_user.bat.
Alternatively, you can run the WORK/update_user.sh from the WSL command line.

# The output files
The following output files will be generated after a successful run:

* **spro_updated_users.csv**  The list of updated user entries. Please import this CSV file to School Pro to complete the user update process.
* **spro_new_users_only.csv**  This file is for reference purpose only. Contains only the new user entries that have been created.
* **refout.csv** This file is for reference purpose only. Contains verbose/informative data that show how each user entry has been modified/created.
