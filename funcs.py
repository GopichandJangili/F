
import pandas as pd,sqlite3,os,numpy as np,itertools
from functools import reduce
from docx import Document
import xlrd,json

def process_sp(dbname):
    if os.path.exists(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')):
        with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')) as json_file:
         json_sort = json.load(json_file)
        sheets_needed=list(json_sort.keys()  )
   
    else:
          sheets_needed,json_sort=None 
    sheets_needed=sheets_needed if sheets_needed is not None else f    
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    template=pd.read_sql('select * from template',conn).drop('index',axis=1)

    template['sheets']=template.apply(lambda x:','.join(list(set([x.split('.')[0] for x in list(itertools.chain.from_iterable([[par.split('}')[0] for par in x[i].split('{') if '}' in par] for i in ['Business_Rule','String_Name']]))]))),axis=1)
 
   
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']

    if f!=[]:
     statement=[]
   
     for db in sheets_needed: 
 
      
      df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)
        
      if json_sort:
       df=df.sort_values(json_sort[db])
      for index,row in df.iterrows():
       template_df=template.loc[template['sheets']==db].reset_index(drop=True)
   
       for i,r in template_df.iterrows():
        a=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.Business_Rule.split('{') if '}' in par]
        b=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.String_Name.split('{') if '}' in par]
        
        for a_i in a:
         
          r.Business_Rule=r.Business_Rule.replace(a_i[0],a_i[1]).replace(a_i[1],str(df.loc[index,a_i[2]]))
        for b_i in b:  
          r.String_Name=r.String_Name.replace(b_i[0],b_i[1]).replace(b_i[1],str(df.loc[index,b_i[2]]))

        if eval(r.Business_Rule):
      
          statement.append(r.String_Name.replace('\r', '').replace('\n', ''))
          
    with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as json_file:
         json_proj = json.load(json_file)  
    d={'#and':' and ','#full':'. ','#com':', ','#sem':'; '}  
    
    if json_proj['rowbinder'] in d.keys() and len(statement)>1:
     statement=f"{d[json_proj['rowbinder']]}".join(statement) 
    elif json_proj['rowbinder'] and len(statement)>1 :
      statement=','.join([a for a in statement if a!=statement[-1]])+' and '+statement[-1]  
    else:
      statement=','.join(statement)
      
    statement=statement+'.' if not statement.endswith('.') else statement
 
     
    return statement
     

def process_pps(dbname):

    if os.path.exists(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')):
        with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')) as json_file:
         json_sort = json.load(json_file)
        sheets_needed=list(json_sort.keys())
        
    else:
          sheets_needed,json_sort=None 
    sheets_needed=sheets_needed if sheets_needed is not None else f 
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    template=pd.read_sql('select * from template',conn).drop('index',axis=1)

    template['sheets']=template.apply(lambda x:','.join(list(set([x.split('.')[0] for x in list(itertools.chain.from_iterable([[par.split('}')[0] for par in x[i].split('{') if '}' in par] for i in ['Business_Rule','String_Name']]))]))),axis=1)
  
   
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']
    
    if f!=[]:
     statement=[]
   
     for db in sheets_needed:  
 
     
      df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)

      if json_sort:

       df=df.sort_values(json_sort[db])

      state=[]
      for index,row in df.iterrows():
       template_df=template.loc[template['sheets']==db].reset_index(drop=True)
  
       for i,r in template_df.iterrows():
        a=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.Business_Rule.split('{') if '}' in par]
        b=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.String_Name.split('{') if '}' in par]
        
        for a_i in a:
         
          r.Business_Rule=r.Business_Rule.replace(a_i[0],a_i[1]).replace(a_i[1],str(df.loc[index,a_i[2]]))
        for b_i in b:  
          r.String_Name=r.String_Name.replace(b_i[0],b_i[1]).replace(b_i[1],str(df.loc[index,b_i[2]]))
    
        if eval(r.Business_Rule):
      
          state.append(r.String_Name.replace('\r', '').replace('\n', ''))
       
      
      with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as json_file:
         json_proj = json.load(json_file)  
      d={'#and':' and ','#full':'. ','#com':', ','#sem':'; '}    
    
      if json_proj['rowbinder'] in d.keys() and len(state)>1 :
        state=f"{d[json_proj['rowbinder']]}".join(state) 
        
      elif json_proj['rowbinder']=='#comand' and len(state)>1:
        state=','.join([a for a in state if a!=state[-1]])+' and '+state[-1]  
      else:
       state=''.join(state)      
      state=state+'.' if not state.endswith('.') and state!='' else state
      
      
      if state!='':
       statement.append(state)
      
 

    
    return statement
     



def process_ppr(dbname):
    if os.path.exists(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')):
        with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')) as json_file:
         json_sort = json.load(json_file)
        sheets_needed=list(json_sort.keys()  )
   
    else:
          sheets_needed,json_sort=None 
    sheets_needed=sheets_needed if sheets_needed is not None else f    
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    template=pd.read_sql('select * from template',conn).drop('index',axis=1)
 
        
    template['sheets']=template.apply(lambda x:','.join(list(set([x.split('.')[0] for x in list(itertools.chain.from_iterable([[par.split('}')[0] for par in x[i].split('{') if '}' in par] for i in ['Business_Rule','String_Name']]))]))),axis=1)

   
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']

    if f!=[]:
     statement=[]
     for db in sheets_needed:  
      df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)
      if json_sort:
    
       df=df.sort_values(json_sort[db])
       
      for index,row in df.iterrows():
       template_df=template.loc[template['sheets']==db].reset_index(drop=True)
       if  not template_df.empty:
   
        state=[]
        for i,r in template_df.iterrows():
         a=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.Business_Rule.split('{') if '}' in par]
         b=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.String_Name.split('{') if '}' in par]
          
         for a_i in a:
         
          r.Business_Rule=r.Business_Rule.replace(a_i[0],a_i[1]).replace(a_i[1],str(df.loc[index,a_i[2]]))
         for b_i in b:  
          r.String_Name=r.String_Name.replace(b_i[0],b_i[1]).replace(b_i[1],str(df.loc[index,b_i[2]]))
    
         if eval(r.Business_Rule):
      
          state.append(r.String_Name.replace('\r', '').replace('\n', ''))
        
        with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as json_file:
         json_proj = json.load(json_file)  
        d={'#and':' and ','#full':'. ','#com':', ','#sem':'; '}    
        if json_proj['rowbinder'] in d.keys() and len(state)>1:
        
         state=f"{d[json_proj['rowbinder']]}".join(state) 
        
        elif json_proj['rowbinder']=='#comand' and len(state)>1:
   
         state=','.join([a for a in state if a!=state[-1]])+' and '+state[-1]  
        else:
         state=''.join(state)
         
        state=state+'.' if not state.endswith('.') else state
    
        statement.append(state)
      

    
    return statement
   
   

def process_ppbr(dbname):
    if os.path.exists(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')):
        with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')) as json_file:
         json_sort = json.load(json_file)
        sheets_needed=list(json_sort.keys()  )
   
    else:
          sheets_needed,json_sort=None 
    sheets_needed=sheets_needed if sheets_needed is not None else f    
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    template=pd.read_sql('select * from template',conn).drop('index',axis=1)
 
        
    template['sheets']=template.apply(lambda x:','.join(list(set([x.split('.')[0] for x in list(itertools.chain.from_iterable([[par.split('}')[0] for par in x[i].split('{') if '}' in par] for i in ['Business_Rule','String_Name']]))]))),axis=1)

   
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']

    if f!=[]:
     statement=[]
     for db in sheets_needed:  
      df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)
      if json_sort:
       df=df.sort_values(json_sort[db])
      for index,row in df.iterrows():
       template_df=template.loc[template['sheets']==db].reset_index(drop=True)
       if  not template_df.empty:
   
   
        for i,r in template_df.iterrows():
         a=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.Business_Rule.split('{') if '}' in par]
         b=[('{'+par.split('}')[0]+'}','{'+par.split('}')[0].split('.')[1]+'}',par.split('}')[0].split('.')[1]) for par in r.String_Name.split('{') if '}' in par]
        
         for a_i in a:
         
          r.Business_Rule=r.Business_Rule.replace(a_i[0],a_i[1]).replace(a_i[1],str(df.loc[index,a_i[2]]))
         for b_i in b:  
          r.String_Name=r.String_Name.replace(b_i[0],b_i[1]).replace(b_i[1],str(df.loc[index,b_i[2]]))
    
         if eval(r.Business_Rule):
          state=r.String_Name.replace('\r', '').replace('\n', '')
          state=state+'.' if not state.endswith('.') else state
    
          statement.append(state)
        
        
      

 
    return statement
   
   
def controller(dbname): 

 word_dict=[]
 conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
 df=pd.read_sql('select * from template',conn).drop('index',axis=1)
 f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]

 if not df.empty and len(f)>1 :
   

      with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as f:
        func=json.load(f)['key']
 
      if func=='#sp':

       word_dict.append(process_sp(dbname))
      elif func=='#pps':
        word_dict.extend(process_pps(dbname)) 
      elif func=='#ppr':
        word_dict.extend(process_ppr(dbname))
      else:
       word_dict.extend(process_ppbr(dbname))
         
      df=pd.DataFrame(word_dict)
   
      if df.empty:
        df=pd.DataFrame(columns=['vals'])
      else:
        df.columns=['vals']
    
      df.to_csv(os.path.join(os.getcwd(),'Data',dbname+'_content.csv'),index=False)  
 else:
        df=pd.DataFrame(columns=['vals'])
        df.to_csv(os.path.join(os.getcwd(),'Data',dbname+'_content.csv'),index=False)
       


def db_store(dbname,*args,**kwargs):
 with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json')) as json_file:
                 proj_json = json.load(json_file)
 filename=kwargs.get('filename',None)
 df=kwargs.get('dataframe',None)
 uploads=kwargs.get('uploads',None)
 conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
 
 f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]

 f=[x for x in f if x!='template']


 if uploads=='input':
 
  for i in f:
   conn.execute(f"drop table {i}")
 
 if filename is not None and df is None:


 

  if uploads =='input':
   if '.csv' in filename: 
    df=pd.read_csv(os.path.join(os.getcwd(),'temp',filename),encoding='ISO-8859-1')
     
    
    df.to_sql('csv',conn,if_exists='replace')
    os.remove(os.path.join(os.getcwd(),'temp',filename))
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']
    b = {}
   
    if f!=[]:
         for db in f:  
          df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)

          b[db] = df.columns.tolist()
      
         
         
         sheet = [(k, v) for k, v in b.items()]
    
   if '.xlsx' in filename:

    xls = xlrd.open_workbook(os.path.join(os.getcwd(),'temp',filename), on_demand=True)  

    for i in xls.sheet_names():
     df=pd.read_excel(os.path.join(os.getcwd(),'temp',filename),sheet_name=i)
     #df=df.astype(str)
     df.to_sql(i,conn,if_exists='replace')
    os.remove(os.path.join(os.getcwd(),'temp',filename))
    conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
    f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]
    f=[x for x in f if x!='template']
    b = {}
   
    if f!=[]:
         for db in f:  
          df=pd.read_sql(f'select * from {db}',conn).drop('index',axis=1)

          b[db] = df.columns.tolist()
      
         
         
         sheet = dict([(k, v) for k, v in b.items()])
         with open(os.path.join(os.getcwd(),'Data',dbname+'_sort.json'),'w') as json_file:
            json.dump(sheet, json_file)
   with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as json_file:
      json_decoded = json.load(json_file) 
   json_decoded['datadefinition']=filename
   with open(os.path.join(os.getcwd(),'Data',dbname+'.json'),'w') as json_file:
            json.dump(json_decoded, json_file)
  if uploads =='template':

   if '.csv' in filename:
    df=pd.read_csv(os.path.join(os.getcwd(),'temp',filename),encoding='ISO-8859-1')
   if '.xlsx' in filename:
   
       
      df=pd.read_excel(os.path.join(os.getcwd(),'temp',filename))
   
   x= [x.lower() for x in df.columns]
   x.sort()
   
   if x==['business_rule','sno','string_name']:
      
    
      df.to_sql(uploads,conn,if_exists='replace')
      os.remove(os.path.join(os.getcwd(),'temp',filename))
     
      with open(os.path.join(os.getcwd(),'Data',dbname+'.json')) as json_file:
          json_decoded = json.load(json_file) 
      json_decoded['datatemplate']=filename
      with open(os.path.join(os.getcwd(),'Data',dbname+'.json'),'w') as json_file:
            json.dump(json_decoded, json_file)    


 if filename is  None and df is not None:

   df['Sno']=np.arange(len(df))+1
   df.to_sql('template',conn,if_exists='replace')




def db_read(dbname,tablename,*args,**kwargs):
 sheets_needed=kwargs.get('sheets_needed',None)

 conn = sqlite3.connect(os.path.join(os.getcwd(),"Data",dbname+'.db'))
 
 f=[','.join(x) for x in conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()]

 if 'template'  in f and tablename=='template':
 
  df=pd.read_sql(f'select * from {tablename}',conn).drop('index',axis=1)

  return df
 elif  'template' not in f and tablename=='template':
  df=pd.DataFrame(columns=["Sno","Business_Rule","String_Name"])
  df.to_sql('template',conn,if_exists='replace')
  return df
 
 
 else:

  f=[x for x in f if x!='template']

  l=[]

  if len(f)!=0:
  
     sheets_needed=sheets_needed if sheets_needed is not None else f   
     
     for x in sheets_needed:
   
      cur=conn.execute(f"select * from {x}")
      l.extend([x+'.'+a[0] for a in cur.description if a[0]!='index'])   
     
   
     return l

  else:
   
   return []  
   
