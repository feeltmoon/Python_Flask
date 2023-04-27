from flask import Flask, render_template, request
import os
import pandas as pd

# create a Flask for our app
app = Flask(__name__)

# create a route for this app
# this route needs to be linked by a function as below def index()
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        fruit = request.form['fruit']
        folder_name = request.form['folder_name']
        path = os.path.join(os.getcwd(), folder_name, fruit)
        if request.form['submit'] == 'Create Folder':
            if os.path.isdir(path):
                return 'Foder already exists'
            else:    
                os.makedirs(path)
                sub_dir = 'Source Files'
                sub_folder = os.path.join(path,sub_dir)
                os.mkdir(sub_folder)
                return 'Folder created'   
        elif request.form['submit'] == 'Generate Reports':
            if fruit == 'apple':
                if os.path.isdir(path):
                    
                    sub_dir_srcfiles = os.path.join(path,'Source Files')
                    
                    if sub_dir_srcfiles == "":
                        return 'Source Files Folder is not existed.'
                        exit()
                    
                    file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'
                                        
                    def find_file(filename, path):
                        for root, dirs, files in os.walk(path):
                            if filename in files:
                                return os.path.join(root,filename)
                    
                    file_path_sugg = find_file(file_name_sugg,sub_dir_srcfiles)
                    
                    if file_path_sugg == "":
                        return 'suggestion file is not uploaded.'
                        exit()
                    # read df1 from suggestion
                    df1 = pd.read_excel(file_path_sugg,sheet_name='Live Contact List - Other',header=1)
                    
                    df1['Role'] = df1['Role'].astype(str)
                    df1['Role'] = df1['Role'].apply(lambda x: x.split('/')).explode().reset_index(drop=True)        
                    df1.to_excel(path + '\\df1_debug.xlsx')
                    # df1 = df1.apply(lambda x: x.str.split('/').explode()).reset_index()
                    
                                     
                    df1['Role'] = df1['Role'].str.lstrip()
                    df1['Role'] = df1['Role'].str.rstrip()
                    # debug_output:
                    df1.to_excel(path + '\\df1_debug.xlsx')
                    # read df2 from suggestion
                    df2 = pd.read_excel(file_path_sugg,sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'])
                    # debug_output:
                    df2.to_excel(path + '\\df2_debug.xlsx')
                    
                    # df3
                    file_name_nmlst = 'Name List.xlsx'
                    file_path_nmlst = find_file(file_name_nmlst,sub_dir_srcfiles)                
                    if file_path_nmlst == "":
                        return 'name list file is not uploaded.'
                        exit()
                    df3 = pd.read_excel(file_path_nmlst,sheet_name='名录（按组织）')
                    df3 = df3.rename(columns={'电子邮件地址': 'Email_Source', '职务头衔':'Title'})
                    def GetEmailAddress(x):
                        return x.split('（')[0].strip(' ')
                    df3.loc[:,'Email'] = df3['Email_Source'].astype(str).apply(lambda x: GetEmailAddress(x))
                    df3 = pd.DataFrame(df3,columns=(['Email','Title']))
                    df3 = df3.drop_duplicates(subset='Email')
                    df3.loc[:,'Email_Upper'] = df3.loc[:,'Email']
                    df3.loc[:,'Email_Upper'] = df3.apply(lambda x: x.str.upper())
                    df3 = df3.drop(columns='Email')
                    # debug_output:
                    df3.to_excel(path + '\\df3_debug.xlsx')
                                     
                    for file in os.listdir(path):
#                        if "name list" in file:
#                            df2 = pd.read_excel(os.path.join(path, file),sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'])
#                        else:
#                            return 'contact list file is not uploaded.'
#                            exit()                      
                        
                        if file.endswith('.xlsx') and file.startswith('Quarterly Access Report'):    
                            
                            df = pd.read_excel(os.path.join(path, file),dtype={'Study Environment Site Number': str},header=11)     
                            
                            for col, values in df.items():
                                if 'Unnamed' in col:
                                    df = df.drop(columns=col)
                            
                            def NoNeedReview(x, y):
                                if '@mdsol.com' in x:
                                    return 'no need to review'
                                elif '@Medidata.com' in x:
                                    return 'no need to review'
                                elif '@medidata.com' in x:
                                    return 'no need to review'
                                elif '@3ds.com' in x:
                                    return 'no need to review'
                                elif y == 'Medidata Internal Beigeneclinical_ebr':
                                    return 'no need to review'  
                            df['Assignment'] = df.apply(lambda x: NoNeedReview(x['Email'], x['Platform Role']),axis=1)
                            df_row = df['Assignment'] != 'no need to review'
                            df_flter = df.loc[df_row,:]
                            df_flter = df_flter.drop(columns = ['Assignment'])
                            # output
                            df.to_excel(path + r'\Review_01.xlsx',index=False)
                            
                        
                        
                        
                        # output report for each study
#                        df.to_excel(os.path.join(path,f'result_{file}'),index=False)
                            
                            
                            
                            
                            
                            # Do something with df
                    return 'All files generated'
                else:
                    print('Folder does not exist')
    return render_template('index.html')
if __name__ == '__main__':
    app.run(debug=True)