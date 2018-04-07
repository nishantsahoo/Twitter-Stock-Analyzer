from xlrd import open_workbook
from textblob import TextBlob
import re
import datetime
import json

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


def getSentiment():
    wb = open_workbook('sample.xlsx')
    sentiment_dict = {}

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        rows = []
        values = []

        for row in range(1, number_of_rows):
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(value)
                except ValueError:
                    pass
                finally:
                    values.append(value)

        date_dict = {}
        for i in range(2,number_of_rows*3-1,3):
            date = values[i-1]
        
            if date in date_dict:
                date_dict[date] += repr(values[i])
            else:
                date_dict[date] = repr(values[i])

        for date, tweets in date_dict.items():
            blob = TextBlob(tweets)
            sentiment = blob.sentiment.polarity
            year, month, day = (int(x) for x in date.split('-'))    
            ans = datetime.date(year, month, day)
            day_name = ans.strftime("%A")
            # print("Sentiment for date: " + date + " ( " + day_name + " ) is: " + str(sentiment))
            sentiment_dict[date] = {
                'sentiment': sentiment,
                'day_name': day_name
            }

    return sentiment_dict

def calDiff():
    wb = open_workbook('AMZN.xlsx')
    change_dict = {}
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
    print(sheet.cell(0,0).value)

    for row in range(1,number_of_rows):
        try:
            date = str((sheet.cell(row,0).value))
            one = (float) (sheet.cell(row,1).value)
            two = (float) (sheet.cell(row,2).value)
            
        except ValueError:
            pass

        if((two-one)>0):
            change_dict[date] = 1

        else:
            change_dict[date] = 0
    return change_dict

def mapSentivalToStockval(sentiment_dict,change_dict):
    newMap_dict = {}
    count=0
    avgVal=0
    day = {'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'}
    for dates in sentiment_dict:
        for datec in change_dict:
            if(dates==datec):
                if(sentiment_dict[dates].day_name).equals(day[5]):
                    flag=1
                elif(sentiment_dict[dates].day_name).equals(day[6]):
                    flag=1
                else:
                    flag=0
                if(flag==1):
                    avgVal+=sentiment_dict[dates].sentiment
                    count+=1
                    continue
                if(sentiment_dict[dates].day_name).equals(day[0]):
                    newMap_dict[dates]={
                    'sentiment': (avgVal/count),
                    'stock': change_dict[datec]
                    }
                    avgVal=0
                    count=0
                else:
                    newMap_dict[dates]={
                    'sentiment':sentiment_dict[dates].sentiment,
                    'stock': change_dict[datec]
                    }
    return newMap_dict


def main():
    sentiment_dict = getSentiment()
    print("Sentiment analysis done... writing the sentiment dictionary into a file.")
    print("Sentiment Dictionary -")
    print(json.dumps(sentiment_dict, sort_keys=True, indent=4))
    file = open("sentiment_dictionary.json","w")
    file.write(json.dumps(sentiment_dict, sort_keys=True, indent=4))
    change_dict = calDiff()
    print("Stock value change :")
    print(json.dumps(change_dict, sort_keys=True, indent=4))
    newMap_dict = mapSentivalToStockval(sentiment_dict,change_dict)
    print("Mapped values :")
    print(json.dumps(newMap_dict, sort_keys=True, indent=4))


main()  # call of the main function