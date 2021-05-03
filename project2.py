
import os
import subprocess
import docx2python
from datetime import datetime
from tqdm import tqdm
from functools import cmp_to_key
import xml.etree.ElementTree as ET
from fuzzywuzzy import fuzz

# global variable to check against []
the_empty_list = []
# global var for dict_file
date_of_acceptance = ''
# global var for dict_file
date_flag = False
# global var for dict_file
dict_arr = []

# dict_file is an XML containing laws titles, page numbers, date of acceptance & signatures


class Point:
    def __init__(self):
        self.header = ''
        self.number = ''
        self.content = ''
        self.sub_points = []


class SubPoint:
    def __init__(self, number, content):
        self.number = number
        self.content = content


class Table:
    def __init__(self, matrix):
        self.matrix = matrix


class Part:
    def __init__(self, num, content):
        self.num = num
        self.content = content


class Chapter:
    def __init__(self, num, content):
        self.num = num
        self.content = content


class Section:
    def __init__(self, num, content):
        self.num = num
        self.content = content


class Signature:
    def __init__(self, role, content):
        self.role = role
        self.content = content


class LawHeader:
    def __init__(self, content):
        self.content = content


class Note:
    def __init__(self, content):
        self.content = content


class DateOfAccept:
    def __init__(self, date):
        self.date = date


class Bind:
    def __init__(self, header, date, lst, signs):
        self.header = header
        self.date = date
        self.lst = lst
        self.signs = signs


# print lists in pretty order
def debug_print(arr):
    for x in arr:
        print(x)
        print('-------------')


# print a list of objects representing a law
def debug_print_objects(law):
    for p in law:
        if type(p) is LawHeader:
            print('         Law Title         ')
            print(p.content)
        elif type(p) is Point:
            print('         Point         ')
            print('Number ' + p.number)
            print('Title ' + p.header)
            print('Content ' + p.content)
            print('Sub Points:')
            for x in p.sub_points:
                print('Number ' + x.number)
                print('Content ' + x.content)
                print('--------')
        elif type(p) is Note:
            print('         Note         ')
            print('Content ' + p.content)
        elif type(p) is Signature:
            print('         Signature         ')
            print('Name: ' + p.content)
            print('Role ' + p.role)
        elif type(p) is Chapter or type(p) is Section or type(p) is Part:
            print('         Section/Chapter/Part         ')
            print('Header: ' + p.num)
            print('Content: ' + p.content)
        print('****   *****   *****')


# not in use after installing tqdm (shows progress bar in console)
def progressBar(iterable, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    total = len(iterable)
    # Progress Bar Printing Function
    def printProgressBar (iteration):
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Initial Call
    printProgressBar(0)
    # Update Progress Bar
    for i, item in enumerate(iterable):
        yield item
        printProgressBar(i + 1)
    # Print New Line on Complete
    print()


# check if all elements in lst are also lists
def all_elements_are_lists(lst):
    if lst == the_empty_list:
        return False
    for x in lst:
        if type(x) is list:
            continue
        else:
            return False
    return True


# is a nested list ALL empty ? ex: [[[]]]
def isListEmpty(inList):
    if isinstance(inList, list): # Is a list
        return all( map(isListEmpty, inList) )
    return False # Not a list


# removes duplicates from any list (Generic)
def removeDuplicates(listToSort):
    no_duplicates = []
    for piece in listToSort:
        if piece in no_duplicates:
            continue
        else:
            no_duplicates.append(piece)
    return no_duplicates


# specified for our project, removes duplicate strings and empty strings
def removeInnerDups(lst):
    if all_elements_are_lists(lst):
        tmp = [removeInnerDups(x) for x in lst]
    else:
        tmp = list(dict.fromkeys(lst))
        while tmp.count('') > 0:
            tmp.remove('')
    return tmp

# removes arrays from the sort of [[]] or [] or [[[]]]
def removeEmptyArrs(nested_lists):
    if not all_elements_are_lists(nested_lists):
        return
    for x in nested_lists:
        if isListEmpty(x):
            nested_lists.remove(x)
    for x in nested_lists:
        if x == the_empty_list:
            nested_lists.remove(x)
        removeEmptyArrs(x)


# content is a list of strings, check the 1st char of the 1st string if its a digit
def startsWithNumber(content):
    if len(content) == 0:
        return False
    if len(content[0]) == 0:
        return False
    return content[0][0].isdigit()


# our own slice string function, handles wrong indices
def sliceString(param, start, stop):
    if stop>len(param):
        return ''
    res = ''
    while start < stop:
        res += param[start]
        start = start + 1
    return res



# content is an array of 2 elements containing side header and the content of the point
# p is the Point object we have to fill
def parsePointContent(content, p):
    if content == the_empty_list:
        return the_empty_list
    # here we seperate the number of the point from the first line
    first_line = content[0]
    res = ''
    i = 0
    while first_line[i].isdigit():
        res = res + first_line[i]
        i = i + 1
        if i == len(first_line):
            break
    # save the number of the point
    p.number = res
    # cut the number out of the line
    first_line = sliceString(first_line, i, len(first_line))
    i = 0
    # figure out where the real letters start after dots and quotes and spaces
    if i < len(first_line):
        while first_line[i] == ' ' or first_line[i] == '*' or first_line[i] == '.' or first_line[i] == ',':
            i = i + 1
            if i == len(first_line):
                break
    # cut the line from where the string begins
    first_line = sliceString(first_line, i, len(first_line))
    # pop the number's string
    content.pop(0)
    # connect the first line to the rest of the arr and combine them
    content.insert(0, first_line)
    inSubPointFlag = False
    # we either have an intro only OR into+subpoints OR subpoints as a Point's content
    for line in content:
        if isSubPoint(line):
            inSubPointFlag = True
            p.sub_points.append(parseSubPoint(line))
        elif inSubPointFlag == False:
            p.content = p.content + '\n' + line
        elif inSubPointFlag == True:
            p.sub_points[-1].content = p.sub_points[-1].content + '\n' + line
    return p


# a higher function that parses Point objects
# Points come in arrays of 2: side note & content
def parsePoint(arr):
    p = Point()
    if startsWithNumber(arr[0]):
        p.header = ' '.join(arr[1])
        parsePointContent(arr[0], p)
    else:
        p.header = ' '.join(arr[0])
        parsePointContent(arr[1], p)
    return p


# parses sub points, extracts the number and returns an object
def parseSubPoint(line):
    num = line[1]
    line = sliceString(line, 3, len(line))
    return SubPoint(num, line)


# parses a chapter, extracts the 2nd word (# of the chapter) and builds an object from the rest
def parseChapter(line):
    lst = line.split(' ')
    if len(lst) < 2:
        return Chapter('', '')
    raw_header = lst[0] + ' ' + lst[1]
    # remove the word chapter and the number
    lst.pop(0)
    lst.pop(0)
    return Chapter(raw_header, ' '.join(lst))


# parses a section
def parseSection(line):
    lst = line.split(' ')
    if len(lst) < 2:
        return Section('', '')
    raw_header = lst[0] + ' ' + lst[1]
    # remove the word chapter and the number
    lst.pop(0)
    lst.pop(0)
    return Section(raw_header, ' '.join(lst))

# parses a section
def parsePart(line):
    lst = line.split(' ')
    if len(lst) < 2:
        return Section('', '')
    raw_header = lst[0] + ' ' + lst[1]
    # remove the word chapter and the number
    lst.pop(0)
    lst.pop(0)
    return Part(raw_header, ' '.join(lst))


# parse the signature, item1 XOR item2 has a keyword role in it ex: prime minister, minister of X
def parseSignature(item1, item2):
    if isSignature(item1):
        z = Signature(item1, item2)
    else:
        z = Signature(item2, item1)
    return z

# returns a DateOfAccept object
def parseDateOfAccept():
    d = DateOfAccept(date_of_acceptance)
    return d


# checks if line is subPoint, starts with "(Hebrew letter)"
def isSubPoint(line):
    tmp = sliceString(line, 0, 3)
    if tmp in ['(א)', '(ב)', '(ג)', '(ד)', '(ל)', '(ה)', '(ו)', '(ז)', '(ח)', '(ט)', '(י)', '(כ)', '(ס)']:
        return True
    return False


# checks if an arr is a Point according to analysis on our tools & files
def isPoint(arr):
    if len(arr) != 2:
        return False
    first = startsWithNumber(arr[0])
    second = startsWithNumber(arr[1])
    if first == second:
        return False
    if len(arr[0]) == 1 and isEndOfPage(arr[0][0]):
        return False
    if len(arr[1]) == 1 and isEndOfPage(arr[1][0]):
        return False
    return True


# check if line is Chapter
def isChapter(line):
    item = (line.split(' ', 1))[0]
    if fuzz.ratio(item, 'פרק') > 60:
        return True
    else:
        return False


def isSection(line):
    item = (line.split(' ', 1))[0]
    if fuzz.ratio(item, 'סימן') > 60:
        return True
    else:
        return False


def isPart(line):
    item = (line.split(' ', 1))[0]
    if fuzz.ratio(item, 'חלק') > 60:
        return True
    else:
        return False


# only the last line in every page could have many spaces
def isEndOfPage(line):
    if '             ' in line:
        return True
    return False


# irrelvant info in the end of page
def isIrrelvantEndOfPage(line):
    if 'נתקבל בכנסת' in line:
        return True
    lst = line.split(' ')
    lst_length = len(lst)
    if lst_length < 3:
        return False
    delim = lst[lst_length-2]
    last = lst[lst_length-1]
    if lst[lst_length-3] == 'עמ':
        delim = lst[lst_length-3]
    if fuzz.ratio(delim, """עמ\'""") > 70:
        if (last != '') and (last[0] == '.' or last[0].isdigit() == True):
            return True
    return False


# the length of line seems absurd but its working
def isSignature(line):
    if len(line) > 20:
        return False
    if fuzz.ratio('נשיא המדינה', line) > 70:
        return True
    if fuzz.ratio('ראש הממשלה', line) > 70:
        return True
    if fuzz.ratio('יושב ראש הכנסת', line) > 65:
        return True
    lst = line.split(' ')
    if fuzz.ratio('שר', lst[0]) > 80:
        return True
    return False


# this function runs one time so no every sentence starting with the word "HOOK" is a title, dont worry
def isLawName(line):
    if fuzz.ratio('חוק', (line.split(' '))[0]) > 70:
        return True
    return False


# Stage 1 & 2 & 3 of the pipeline, convert to docx, read the contents, fix the arrays
def prepare_text(file_name, main_path):
    # create the path to the file
    fn = os.path.join(main_path, file_name)
    # convert it
    subprocess.call(['C:\Program Files\LibreOffice\program\soffice.exe', '--headless', '--convert-to', 'docx', fn], shell=True)
    file_name = file_name[:-4] + '.docx'
    # read the file
    content = docx2python.docx2python(file_name, extract_image=False)
    # remove the docx after reading
    os.remove(file_name)
    # take the content with the structure it came with
    content = content.body
    #debug_print(content)
    # remove duplicate nodes in the lists, remove duplicate strings at the lowest level
    no_duplicates = removeInnerDups(removeDuplicates(content))
    # debug_print(no_duplicates)
    # remove empty arrs including the sorts of [[[]]]
    removeEmptyArrs(no_duplicates)
    return no_duplicates


# parses a line, appends a proper object to res list
def parseLine(res, line):
    global date_flag, date_of_acceptance, dict_arr
    if isLawName(line) and isIrrelvantEndOfPage(line) == False:
        res.append(LawHeader(line))
        dict_arr[-1].header = line
        return True
    elif isSection(line):
        res.append(parseSection(line))
        return True
    elif isChapter(line):
        res.append(parseChapter(line))
        return True
    elif isPart(line):
        res.append(parsePart(line))
        return True
    elif isSubPoint(line) and len(res) > 0:
        if type(res[-1]) is Point:
            res[-1].sub_points.append(parseSubPoint(line))
            return True
    elif isEndOfPage(line):
        tmp = line.split(' ')
        if date_flag == False:
            if len(tmp) > 0:
                date_of_acceptance = tmp[-1]
                dict_arr[-1].date = tmp[-1]
                date_flag = True
        if tmp[0].isnumeric():
            dict_arr[-1].lst.append(int(tmp[0]))
        return True
    elif isIrrelvantEndOfPage(line):
        return True
    # continuation of a subpoint
    elif len(res) > 0 :
        if (type(res[-1]) is Point) and (res[-1].sub_points != the_empty_list):
            res[-1].sub_points[-1].content = res[-1].sub_points[-1].content + '\n' + line
            return True
    elif res != the_empty_list and len(line) > 30:
        res[-1].content = res[-1].content + '\n' + line
        return True
    elif res != the_empty_list:
        if type(res[-1]) is Note:
            res[-1].content = res[-1].content + '\n' + line
            return True
    else:
        return False


def parse(body):
    global date_flag, dict_arr
    results = []
    # need this variable to skip iterations in FOR loops in python
    control_switch = False
    # dict_arr is a dicitionary for all laws, for every law we append another Bind item
    dict_arr.append(Bind('', '', [], []))

    for x in body:
        for block in x:
            if isPoint(block):
                results.append(parsePoint(block))
            else:
                for part in block:
                    for i in range(len(part)):
                        if control_switch == True:
                            control_switch = False
                            continue
                        if (i+1) < len(part):
                            if isSignature(part[i]) or isSignature(part[i+1]):
                                tmp = parseSignature(part[i], part[i+1])
                                dict_arr[-1].signs.append(tmp)
                                results.append(tmp)
                                control_switch = True

                        if control_switch == False and parseLine(results, part[i]):
                            continue
                        else:
                            results.append(Note(''))
                            parseLine(results, part[i])
    results.append(parseDateOfAccept())
    date_flag = False
    return results


# not in use, still needs testing, sorts the points in order and makes sure no
# weird ups and downs in the rising series of numbers as points numbers
def fix_law(lego):
    first_index = 1
    flag = True
    result = []
    for item in lego:
        if type(item) is Point and flag:
            first_index = int(item.number)
            last_point = item
            flag = False
            result.append(item)
            continue
        elif type(item) is Point:
            if int(item.number) == (first_index + 1):
                first_index = first_index + 1
                last_point = item
                result.append(item)
                continue
            else:
                last_point.content = last_point.content + '\n' + item.header + ' ' + item.number + '\n' + item.content
                for x in item.sub_points:
                    last_point.content = last_point.content + '\n' + x.number + ' ' + x.content
        elif type(item) is Note and item.content == '':
            continue
        else:
            result.append(item)
    return result


# creates an XML file from lego (list of objects as one LAW) and a filename
def create_xml(lego, file_name):
    root = ET.Element('body')
    last_big_item = root
    part_stack = [root]
    chapter_stack = [root]
    law_name_flag = False

    for item in lego:
        if type(item) is LawHeader and law_name_flag == False:
            title = ET.SubElement(root, 'title')
            content = ET.SubElement(title, 'content')
            p = ET.SubElement(content, 'p')
            p.text = item.content
            law_name_flag = True
        elif type(item) is Chapter and item.content != '':
            chapter = ET.SubElement(part_stack[-1], 'chapter')
            num = ET.SubElement(chapter, 'num')
            num.text = item.num
            title = ET.SubElement(chapter, 'title')
            content = ET.SubElement(title, 'content')
            p = ET.SubElement(content, 'p')
            p.text = item.content

            chapter_stack.append(chapter)
            last_big_item = chapter
        elif type(item) is Section and item.content != '':
            section = ET.SubElement(chapter_stack[-1], 'section')
            num = ET.SubElement(section, 'num')
            num.text = item.num
            title = ET.SubElement(section, 'title')
            content = ET.SubElement(title, 'content')
            p = ET.SubElement(content, 'p')
            p.text = item.content

            last_big_item = section
        elif type(item) is Part and item.content != '':
            part = ET.SubElement(root, 'part')
            num = ET.SubElement(part, 'num')
            num.text = item.num
            title = ET.SubElement(part, 'title')
            content = ET.SubElement(title, 'content')
            p = ET.SubElement(content, 'p')
            p.text = item.content

            part_stack.append(part)
            last_big_item = part
        elif type(item) is Note and item.content != '':
            note = ET.SubElement(last_big_item, 'note')
            note.text = item.content
        elif type(item) is Signature:
            sig = ET.SubElement(root, 'signature')
            name = ET.SubElement(sig, 'name')
            name.text = item.content
            role = ET.SubElement(sig, 'role')
            role.text = item.role
        elif type(item) is Point:
            point = ET.SubElement(last_big_item, 'point')
            heading = ET.SubElement(point, 'heading')
            auth = ET.SubElement(heading, 'authorialNote', {'placement': 'side'})
            p = ET.SubElement(auth, 'p')
            p.text = item.header
            num = ET.SubElement(point, 'num')
            num.text = item.number

            if item.sub_points == the_empty_list:
                content = ET.SubElement(point, 'content')
                p = ET.SubElement(content, 'p')
                p.text = item.content
            else:
                lst = ET.SubElement(point, 'list')
                intro = ET.SubElement(lst, 'intro')
                p = ET.SubElement(intro, 'p')
                p.text = item.content

                for x in item.sub_points:
                    sub = ET.SubElement(lst, 'subPoint')
                    num = ET.SubElement(sub, 'num')
                    num.text = x.number
                    content = ET.SubElement(sub, 'content')
                    p = ET.SubElement(content, 'p')
                    p.text = x.content
        elif type(item) is DateOfAccept:
            d = ET.SubElement(root, 'dateOfAcceptance')
            d.text = item.date
        else:
            continue

    file_name = file_name[:-4] + '.xml'
    with open(file_name, 'wb') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" ?>\n'.encode('utf8'))
        tree = ET.ElementTree(root)
        tree.write(f, encoding="utf-8")


# comprator for sorting the dict_file
def dateCompare(x, y):
    fst = datetime.strptime(x.date, '%d.%m.%Y')
    snd = datetime.strptime(y.date, '%d.%m.%Y')
    if fst.year > snd.year:
        return 1
    elif fst.year < snd.year:
        return -1
    elif fst.month > snd.month:
        return 1
    elif fst.month < snd.month:
        return -1
    elif fst.day > snd.day:
        return 1
    elif fst.day < snd.day:
        return -1
    else:
        return 0


# creates dict_file, sorted according to dates
def create_dict():
    root = ET.Element('dictionary')
    no_empty_dict = []
    for x in dict_arr:
        if len(x.lst) != 0 and len(x.date) == 10:
            no_empty_dict.append(x)
    new_op = cmp_to_key(dateCompare)
    no_empty_dict.sort(key=new_op, reverse=False)
    for b in no_empty_dict:
        tmp = ''
        law = ET.SubElement(root, 'law')
        ET.SubElement(law, 'header').text = b.header
        ET.SubElement(law, 'date').text = b.date
        for x in b.lst:
            tmp = tmp + str(x) + ', '
        ET.SubElement(law, 'pages').text = tmp
        for x in b.signs:
            t = ET.SubElement(law, 'signature')
            ET.SubElement(t, 'name').text = x.content
            ET.SubElement(t, 'role').text = x.role
    with open('dict_file.xml', 'wb') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" ?>\n'.encode('utf8'))
        tree = ET.ElementTree(root)
        tree.write(f, encoding="utf-8")


# TODO: upgrades timelines for each law, connecting laws references, ordering HAGDAROT sections
def main():
    # input folder
    main_path = os.path.join(os.getcwd(), 'PDF_DOC')
    # output folder
    out_folder = os.path.join(os.getcwd(), 'output')
    # list of files in input
    list_of_files = os.listdir(main_path)
    # redirect working directory to output folder
    os.chdir(out_folder)
    # wrapping the list in tqdm library for showing the progress bar
    for fn in tqdm(list_of_files):
        # if filename ends with DOC and there is no XML for it, run the pipeline
        if fn.endswith(".DOC"):
            if (fn[:-4]+'.xml') in os.listdir(out_folder):
                continue
            # if fn[:-4] == '146894':
            organized = prepare_text(fn, main_path)
            law = parse(organized)
            create_xml(law, fn)
    # create dictionary for parsed laws, sorted to dates of acceptance
    create_dict()


if __name__ == "__main__":
    main()
