#!/bin/sh

# set locale to UTF-8 
export LANG=ja_JP.utf8

cd WORK

# delete stale output files
rm -f ../spro_new_users_only.csv ../refout.csv ../spro_updated_users.csv


# pre-process input CSVs
# translate from sjis to UTF-8

for I in spro_current_users school_student_list; do
    IFN=../InputCSV/$I.csv
    if [ ! -f $IFN ]; then
	echo '***エラー: 入力ファイル' InputCSV/$I.csv 'が存在しません！***'
	exit 1
    fi
    nkf -c -w80 $IFN  > $I.txt
done


# Execute the python script
if ! ./update_user.py spro_updated_users.txt spro_new_users_only.txt refout.txt spro_current_users.txt school_student_list.txt; then
	exit 1
fi

# pos-process the output CSVs
# translate from UTF-8 to sjis

nkf -s -c spro_updated_users.txt > ../spro_updated_users.csv
nkf -s -c spro_new_users_only.txt > ../spro_new_users_only.csv
nkf -s -c refout.txt         > ../refout.csv



