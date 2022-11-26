
import csv
import io
import codecs

def read_csv(fileName):
  csvDict = {}
  try:
    with open(fileName, 'r') as f:
      csvDict = [*csv.DictReader(f)]
  except Exception as e:
    print (f"Error reading data file '{fileName}': {e}")
    print("Repairing data")
    with open(fileName, 'r') as f:
      numLines = 0
  # Get the column names
      line = f.readline().strip()
      columns = line.split(',')
      numColumns = len(columns)
      print(f"There are {numColumns} columns.")
      # print(f"  {columns}")

      lines = []
      histogramCount = {}
      while True:
        line = f.readline().strip()
        lineElements = line.split(',')
        if not line:
          break
        elif line == '\x00':
            continue
        else:
          numElements = len(lineElements)
          try:
              histogramCount[numElements] += 1
          except KeyError:
              histogramCount[numElements] = 1
          # if numColumns != numElements:
          #   print(f"{numLines}:{numColumns}", end=' ')
          # numLines += 1
          lines.append(line)
      print(f"rows per element count = {histogramCount}")
      print()
      print(numLines)
      print(len(lines))
      # csvDict = [*csv.DictReader(lines)]

  try:
    numItems = len(csvDict)
    if numItems != 0:
      print(f"number of entries: {numItems}")
      print("Printing up to 5 samples of the dataset")
      tenPercent = int(float(numItems) * .1)
    else:
      print(f"No data in the file: '{fileName}'")
    ii = 0
    for i, item in enumerate(csvDict, 1):
# print the first 5 entries of a 10% sample of the data
      if i%tenPercent == 0:
        ii += 1
        if ii > 5:
          break
        else:
          print(f"{i:4}. {item['Last Name']}, {item['First Name']}, {item['Email Address']}")
  except Exception as e:
    print (f"Error reading csv file '{fileName}': {e}")
  return csvDict


if __name__ == "__main__":
  memberDict = read_csv("pqi-etl/data/20221125_members_export_excel_ymx.ymx")
  memberDict = read_csv("pqi-etl/data/20221125_members_export_ymt.ymt")
  # non_memberDict = read_csv("pqi-etl/data/20221123_non_members_export_batch.ymx")
