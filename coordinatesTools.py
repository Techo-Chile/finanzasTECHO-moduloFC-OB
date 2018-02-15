import re
def incrementNumber(number):
  ret = "";
  i = len(number) - 1;
  while True:
    if (i < 0):
      return "1" + ret;
    if number[i] == '9':
      ret = "0" + ret;
      i -= 1
    else:
      ret = chr(ord(number[i]) + 1) + ret;
      break;
  ret = number[0:i] + ret;
  return ret;
def incrementWord(word):
  ret = ""
  i = len(word) - 1;
  while True:
    if (i < 0):
      return "A" + ret;
    if word[i] == 'Z':
      ret = "A" + ret;
      i -= 1
    else:
      ret = chr(ord(word[i]) + 1) + ret;
      break;
  ret = word[0:i] + ret;
  return ret;
def splitCoordinates(s):
    return filter(None, re.split(r'(\d+)', s))
def translateCoordinates(x, y):
  return "" + chr(x + 64) + chr(y + ord('0'));
def increaseRow(coord):
  coordinates = splitCoordinates(coord);
  coordinates[1] = incrementNumber(coordinates[1]);
  return ''.join(coordinates);
def increaseColumn(coord):
  coordinates = splitCoordinates(coord);
  coordinates[0] = incrementWord(coordinates[0]);
  return ''.join(coordinates);