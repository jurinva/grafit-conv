#!/bin/bash

mkdir conv

f=''
for I in `ls *.GFT | sort -n`; do
  filename=`echo $I | cut -d"." -f1`
  > ./conv/$filename.csv
  date=`date -r ./$I +%F`
  desc1=`iconv -fcp1251 -tutf8 ./$I | head -1 | cut -d";" -f1 | sed 's/: /:/g'`
  desc2=`iconv -fcp1251 -tutf8 ./$I | head -1 | cut -d";" -f2 `
  if [ "$desc1" != "$desc2" ]; then echo $desc1 $desc2 >> ./conv/$filename.csv
    else echo $desc2 >> ./conv/$filename.csv
  fi
  echo $date >> ./conv/$filename.csv
  echo "V0;S0" >> ./conv/$filename.csv
  sed -n '2,1p' ./$I | sed 's/ /; /g' | sed 's/\./,/g' >> ./conv/$filename.csv
  echo "diameter(mm);height(mm);weight(g);effort(kg)" >> ./conv/$filename.csv
  sed -n '4,544p' ./$I | sed 's/ /; /g' | sed 's/\./,/g' >> ./conv/$filename.csv
  soffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --infilter="csv:59,34,UTF8" ./conv/$filename.csv --outdir ./conv
  f="$f $filename.xlsx"
done

cd ./conv
ssconvert -S -M ./0.xlsx $f