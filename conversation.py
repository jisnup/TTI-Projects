import pandas as pd

df = pd.read_excel('C:/Users/Jisnu/Desktop/jisnu/Data/conversation_length.xlsx')

Auto = ['QSA 2' , 'QSA T3', 'QSA 3','EQA 3', 'EQA 2', 'EQA'
,'Adharmonics', 'Connect Quote','CIF - Auto','Prestige 2', 'Prestige 3','A2-S' , 'A3' ,'A2' , 'Inside - A2',
'TLC', 'TLC 3', 'Everquote(C)','Datalot', 'MF - Florida', 'MF 2 - Florida',
'QuoteStorm', 'Quotestorm 2', 'QuoteStorm EXD', 'Quotestorm EX','Quotestorm 3', 'Quotestorm T3'
]

Home = ['EQ Home', 'TLC Home', 'CIA - Home', 'EQA Home']
#'Hometown'

FL = ['Fl Sales','Fl Trainee - Development', 'Fl Trainee - Acquisition', 'Fl Sales Spanish']
GA = ['GA - Sales','Ga - Trainees']
talk = ['Conversation']
no_talk = ['Unknown','Voicemail','Transfer']
df['Lead Source'] = df['Lead Source'].replace(Auto, 'Auto')
df['Lead Source'] = df['Lead Source'].replace(Home, 'Home')
df['Group'] = df['Group'].replace(FL, 'FL')
df['Group'] = df['Group'].replace(GA, 'GA')
df['Result'] = df['Result'].replace(talk, 1)
df['Result'] = df['Result'].replace(no_talk, 0)
df['Result'] = pd.to_numeric(df['Result'])
df['Call Duration'] = df['Call Duration'].map(lambda x: str(x)[:-3])
df['Call Duration'] = df['Call Duration'].map(lambda x: str(x).replace(":",'.'))
df['Call Duration'] = pd.to_numeric(df['Call Duration'])
df.drop(['Time'], axis = 1, inplace = True)
Grouped = df.groupby(['Group','Lead Source']).agg({'Call Duration':['count','sum', 'mean'], 'Result':['sum']}).fillna(0)

#Grouped = df.pivot_table(index='Group',columns='Lead Source',aggfunc= sum)
print(Grouped)
