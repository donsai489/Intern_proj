import nltk,xlrd,codecs,datetime,csv,io,autocorrect
missing=["miss","missing","did","miss nozzle","nosel","missing nozzle","missed","not found","nozzle"]
damage=["faulty","fault","leaking","leak","leakage","damage","broke","broken","damaged","bad","disappoint","worst","damaging"]
good=["good","gr8","worth","useful","perfectly","reasonable","works","very","easy","better","gud","use","liked","smooth","value","like","excellent","best","wonder","nice","satisfied","like","very good","perfect","great","wow","wonderful","fantastic"]
fake=["fake","mistake"]
sticky=["sticky","stick"]
from autocorrect import Speller
sp=Speller(lang='en')
g_count=0
m_count=0
d_count=0
arf=[]
s_count=0
f_count=0
mi_count=0
mis_count=0
def tok(bb):
    a=[]
    a=nltk.word_tokenize(bb[3])
    return a
from nltk.corpus import stopwords
def m(c):
    return any(item in missing for item in c)
def d(c):
    return any(item in damage for item in c)
def g(c):
    return any(item in good for item in c)
def f(c):
    return any(item in fake for item in c)
def stt(c):
    return any(item in sticky for item in c)
def starsa(bb):
    if(bb[1]!="no_of_stars"):
        if(int(float(bb[1]))==5 or int(float(bb[1]))==4):
            return "g"
        elif(int(float(bb[1]))==1 or int(float(bb[1]))==2 or int(float(bb[1]))==3):
            return "b"
from nltk.stem import PorterStemmer    
ps=PorterStemmer()
def check(a,rel):
    global m_count
    global g_count
    global d_count
    global f_count
    global mi_count
    global s_count
    sw=set(stopwords.words('english'))
    a=[w for w in a if not w in sw]
    b=[i.lower() for i in a]
    c=[sp(i) for i in b]
    
    for r in c:
        r=ps.stem(r)
    if(m(c)==True):
        m_count+=1
    elif(m(c)==False and g(c)==True and d(c)==False and stt(c)==False):
        g_count+=1
    elif(g(c)==False and d(c)==True and stt(c)==False):
        d_count+=1
    elif(f(c)==True):
        f_count+=1
    elif(d(c)==True or m(c)==True or stt(c)==True):
        s_count+=1
    else:
        mi_count+=1
        arf.append(c)  
 
l=(r'C:\Users\saiiyer23\Downloads\test.xlsx')
rel=""
w=xlrd.open_workbook(l)
s=w.sheet_by_index(0)
s.cell_value(0,0)
for i in range(s.nrows):
    bb=s.row_values(i)
    re1=starsa(bb)
    a=tok(bb)
    check(a,rel)

for i in arf:
    if(g(i)==True):
        g_count+=1
    elif(stt(i)==True):
        s_count+=1
    elif(d(i)==True):
        d_count+=1
    else:
        mis_count+=1
print("MISSING-",m_count)
print("GOOD-",g_count)
print("STICKY-",s_count)
print("DAMAGED-",d_count)
print("FAKE-",f_count)
print("MIX_OF_REACTION-",mis_count)
import xlwt
from xlwt import Workbook
wb=Workbook()
she=wb.add_sheet('Sheet 1')
pj=[m_count,g_count,s_count,d_count,f_count,mis_count]
jp=["MISSING NOZZLE","GOOD CONDITION","STICKY & LEAKAGE","DAMAGE","FAKE","MIX_OF_REACTION"]
for i in range(len(jp)):
    she.write(i,0,jp[i])
    she.write(i,1,pj[i])
wb.save("xlwt example.xls")


