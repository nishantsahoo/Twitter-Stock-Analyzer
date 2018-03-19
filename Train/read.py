from xlrd import open_workbook
from textblob import TextBlob
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

wb = open_workbook('sample.xlsx')
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    rows = []
    values = []
    tweet_string = []
    for row in range(22788, 24818):
        for col in range(number_of_columns):
            value  = (sheet.cell(row,col).value)
            if col==2:
                tweet_string += (preprocess(value))
            try:
                value = str(value)
            except ValueError:
                pass
            finally:
                values.append(repr(value))

    cnt = 1
    avg = 0
    text = ""
    for i in range(2,2200,3):
        text += repr(values[i])
        # print text

    blob = TextBlob(text)
    print("Blob sentiment: " + str(blob.sentiment.polarity))
    #    avg += blob.sentiment.polarity
    #    cnt += 1
