# -*- coding: utf-8 -*-

import os
import win32com.client as win32

# create list of files (exclude temporarly word files (starts with '~$')
def walk(dir):
  for name in os.listdir(dir):
    path = os.path.join(dir, name)
    fileslist = open('fileslist.txt', 'a+')
    if os.path.isfile(path):
        if name.startswith('~$'):
            os.remove(path)
        else:
            fileslist.write(path)
            fileslist.write('\n')
    else:
        walk(path)


# extract statistics from .doc files
def wrdreader(link):
  word = win32.gencache.EnsureDispatch('Word.Application')
  word1 = word.Documents.Open(link)
  word1.Visible = False
  word1text = word1.Content
  # counts.txt - a list of lines, where a line have the next format: file with filename + number of word statistic
  filescount = open('counts.txt', 'a+')
  # numbers.txt - a line of numbers, separated by a space
  numberscount = open('numbers.txt', 'a+')
  # Below we are used .ComputeStatistics() method with '5' as an argument. The next variations can be used: '0' - StatisticWords, '1' - StatisticLines', '2' - StatisticPages, '3' - StatisticCharacters, '4' - StatisticParagraphs, '5' - StatisticCharactersWithSpaces
  filescount.write(link + ' ' + str(word1text.ComputeStatistics(5)) + '\n')
  numberscount.write(str(word1text.ComputeStatistics(5))+' ')
  print(link)
  print(word1text.ComputeStatistics(5))
  filescount.close
  numberscount.close
  word1.Close()
  
  

def countr():
  numberscount = open('numbers.txt')
  arrnumbers = numberscount.read().strip()
  summ=0
  for i in arrnumbers.split(' '):
    if type(int(i)) == int:
      summ+=int(i)   
  numberscount.close
  numberscount = open('numbers.txt', 'a+')
  numberscount.write('\nTotal: '+str(summ))
  print(summ)

  
def main():
# for the correct operation of the program need to specify the directory of word files (program can correctly process, including the situation when files are placed in subfolders)
  path = 'c:\counter\wordfiles'
# remove temp files for earlier execution of this app
  for file in ['counts.txt', 'fileslist.txt', 'numbers.txt']:
    if os.path.isfile(file): 
      os.remove(file)
  walk(path)
  fileslist = open('fileslist.txt', 'r')
  links_to_all_files = fileslist.read()
  links = links_to_all_files.split('\n')
  for link in links:
    if link != '':
      wrdreader(link)
  countr()


if __name__ == '__main__':
    main()


# http://www.blog.pythonlibrary.org/2010/07/16/python-and-microsoft-office-using-pywin32/
# https://habrahabr.ru/post/232291/#com
