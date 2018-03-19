from xlrd import open_workbook

import re
 
emoticons_str = r"""
    (?:
        [:=;] # Eyes
        [oO\-]? # Nose (optional)
        [D\)\]\(\]/\\OpP] # Mouth
    )"""
 
regex_str = [
    emoticons_str,
    r'<[^>]+>', # HTML tags
    r'(?:@[\w_]+)', # @-mentions
    r"(?:\#+[\w_]+[\w\'_\-]*[\w_]+)", # hash-tags
    r'http[s]?://(?:[a-z]|[0-9]|[$-_@.&amp;+]|[!*\(\),]|(?:%[0-9a-f][0-9a-f]))+', # URLs
 
    r'(?:(?:\d+,?)+(?:\.?\d+)?)', # numbers
    r"(?:[a-z][a-z'\-_]+[a-z])", # words with - and '
    r'(?:[\w_]+)', # other words
    r'(?:\S)' # anything else
]
    
tokens_re = re.compile(r'('+'|'.join(regex_str)+')', re.VERBOSE | re.IGNORECASE)
emoticon_re = re.compile(r'^'+emoticons_str+'$', re.VERBOSE | re.IGNORECASE)
 
def tokenize(s):
    return tokens_re.findall(s)
 
def preprocess(s, lowercase=False):
    tokens = tokenize(s)
    if lowercase:
        tokens = [token if emoticon_re.search(token) else token.lower() for token in tokens]
    return tokens

class Arm(object):
    def __init__(self, id, date, tweetcontent):
        self.Id=id
        self.Date=date
        self.Tweetcontent=tweetcontent

    def __str__(self):
        return self.Id+' '+self.Date+' '+self.Tweetcontent

wb = open_workbook('sample.xlsx')
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    rows = []
    values = []
    for row in range(1, 10):
        for col in range(number_of_columns):
            value  = (sheet.cell(row,col).value)
            try:
                value = str(value)
            except ValueError:
                pass
            finally:
                print preprocess(value)
                print "next"
                values.append(value)

        print "end line"
        # print values
        print "\n"