IF EXIST out.csv DEL /F out.csv
java -jar JSONTOCSV.jar %1 > out.csv
