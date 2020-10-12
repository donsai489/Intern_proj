import nltk,xlrd,codecs,datetime,csv,io,autocorrect

from nltk import pos_tag
from nltk.corpus import stopwords
from autocorrect import Speller
sp=Speller(lang='en')
g_count=0
d_count=0

from nltk.stem import PorterStemmer    
ps=PorterStemmer()

l=(r'C:\Users\saiiyer23\Downloads\test.xlsx')
rel=""
w=xlrd.open_workbook(l)
s=w.sheet_by_index(0)
s.cell_value(0,0)
import nltk.classify.util
from nltk.classify import NaiveBayesClassifier
from nltk.corpus import movie_reviews
 
def word_feats(words):
    return dict([(word, True) for word in words])

def format_scenetence(sent):
    return({word: True for word in nltk.word_tokenize(sent)})
from nltk.stem import WordNetLemmatizer
lemma=WordNetLemmatizer()

negids = movie_reviews.fileids('neg')
posids = movie_reviews.fileids('pos')
 
negfeats = [(word_feats(movie_reviews.words(fileids=[f])), 'neg') for f in negids]
posfeats = [(word_feats(movie_reviews.words(fileids=[f])), 'pos') for f in posids]
 
negcutoff = int(len(negfeats)*3/4)
poscutoff = int(len(posfeats)*3/4)
 
trainfeats = negfeats[:negcutoff] + posfeats[:poscutoff]
testfeats = negfeats[negcutoff:] + posfeats[poscutoff:]
 
classifier = NaiveBayesClassifier.train(trainfeats)
print ('accuracy:', nltk.classify.util.accuracy(classifier, testfeats))
stemmer = nltk.stem.snowball.SnowballStemmer('english')
default_stopwords = set(nltk.corpus.stopwords.words('english'))
all_stopwords = default_stopwords 
sor=""
n_count=0
for i in range(s.nrows):
    bb=s.row_values(i)
    sor+=bb[3]
    reee=classifier.classify(format_scenetence(bb[3]))
    
    if(reee=="pos"):
        g_count+=1
    elif(reee=="neg"):
        d_count+=1
    else:
        n_count+=1

print("GOOD-",g_count)
print("DAMAGED-",d_count)
disc={}
words = nltk.word_tokenize(sor)
words = [word for word in words if len(word) > 1]
words = [word for word in words if not word.isnumeric()]
words = [word.lower() for word in words]
'''
words = [stemmer.stem(word) for word in words]
'''
words = [word for word in words if word not in all_stopwords]
words=[word for word in words if word.isalpha()]

fdist = nltk.FreqDist(words)
for word, frequency in fdist.most_common(30):
    word=sp(word)
    mn=pos_tag(nltk.word_tokenize(word))
    
    no=[word for word,pos in mn if (pos!='NNPS' and pos!='NNP' and pos!='RBR' and pos!='RBS' and pos!='VB' and pos!='VBG' and pos!='RB' and pos!='NNS' and pos!='VBD' and pos!='VBN' and pos!='VBP' and pos!='VBZ')]

    if(len(no)!=0):
    
        print(u'{};{}'.format(word, frequency))
        disc[word]=frequency 


import xlwt
from xlwt import Workbook
wb=Workbook()
she=wb.add_sheet('Sheet 1')
pj=[g_count,d_count,n_count]
jp=["GOOD","DAMAGE","NEUTRAL"]
for i in range(len(jp)):
    she.write(i,0,jp[i])
    she.write(i,1,pj[i])
wb.save("xlwt example.xlsx")

import xlsxwriter
wb1=xlsxwriter.Workbook('xLsx.xlsx')
wb2=wb1.add_worksheet()
col=0
for key,value in disc.items():
    wb2.write(0,col,key)
    wb2.write(1,col,str(value))
    col+=1
wb1.close()
