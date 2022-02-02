import pywinauto, time,re
from pywinauto.application import Application
from pywinauto import keyboard
def bot(keyword,path):

        
        app = Application().start(cmd_line=r"C:\Program Files\EVC\WinOLS\ols.64Bit.exe")
        #app.WindowSpecification
        #time.sleep(10)
        #app = Application().connect(title_re=".*WinOLS.*")
        app.window(title_re=".*WinOLS.*").wait('visible',timeout=20)

        dlg = app.window(title_re=".*WinOLS.*")
        time.sleep(2)
        count=0
        while True:
            
            
                isset=0
                dlg.Button1.click_input()

                time.sleep(1)
                dlg = app.window(title_re=".*Select.*")
                dlg.child_window(auto_id='StaticWrapper')
                #List2 {UP} Button5 Button1 DESKTOP-KKUI2UF_10000
                dlg.Edit.click_input()
                keyword = keyword.replace(' ','{SPACE}')
                dlg.type_keys(keyword)
                try:
                    cols = dlg.ListView1.columns()
                    ites = dlg.ListView1.items()
                    for i in range(dlg.ListView1.item_count()):
                        item="|".join([ites[i*len(cols)+j].text()for j in range(len(cols))])
                        if ("Exported" in item)==False:
                            isset+=1
                            dlg.ListView1.get_item(i).click_input()
                            break
                        
                except Exception as e:
                    print (e.message, e.args)
                    save = open('logs.txt','a')
                    save.write(e.message, e.args)
                if isset == dlg.ListView1.item_count()-1:
                    break
                dlg.Button2.click_input()
                time.sleep(2)
                
                dlg = app.window(title_re=".*version.*")
                #dlg.print_control_identifiers() Button23
                ver=1
                if (dlg.TreeView.item_count()) >2:
                    ver = input('Entrer le numero de version :')
                    time.sleep(5)
                    
                dlg.Button1.click_input()
                try:
                    
                    dlg = app.window(title_re=".*Checksums.*")
                    time.sleep(2)
                    dlg.OK.click_input()
                except:
                    
                    pass
                
                dlg = app.window(title_re=".*WinOLS.*")

                dlg.Button4.click_input()
                try:
                    dlg = app.window(title_re=".*WinOLS 5.*")
                    dlg.Button2.click_input()
                    time.sleep(1)
                except :
                    pass
                
                dlg = app.window(title_re=".*File.*")

                if(dlg.Button9.get_check_state())==1:
                    dlg.Button9.click_input()
                dlg.OK.click_input()
                eng=app.window(title_re=".*Enregistrer.*")
                time.sleep(1)
                eng.Address_Band_Root.type_keys(r'{F4} ^a {DELETE}')
                #eng.Address_Band_Root.type_keys('{DELETE}')

                #Swap send_keys('^a^c') {DELETE} {ENTER} {SPACE} C:\Program Files\EVC\WinOLS 

                path=path.replace(' ','{SPACE}')
                eng.Address_Band_Root.type_keys(r""+path)
                eng.Address_Band_Root.type_keys('{ENTER}')
                if "(*.*)" in eng.ComboBox1.texts():
                    eng.ComboBox2.click_input()
                else:
                    eng.ComboBox1.click_input()
                keyboard.send_keys(r'^c')
                keyboard.send_keys(r'^+n')
                keyboard.send_keys(r'^v')
                
                #{ENTER}
                keyboard.send_keys(r'{ENTER}')
                keyboard.send_keys(r'{ENTER}')

                eng.Enregistrer.click_input()
                dlg = app.window(title_re=".*WinOLS.*")
                dlg.Button5.click_input()
                dlg = app.window(title_re=".*version.*")
                count=(dlg.TreeView.item_count())
                for i in range(1,int(count)):
                    (dlg.TreeView.get_item([i]).click_input())
                    dlg.OK.click_input()
                    
                    
                    
                    dlg = app.window(title_re=".*WinOLS.*")

                    dlg.Button4.click_input()
                    try:
                        dlg = app.window(title_re=".*WinOLS 5.*")
                        dlg.Button2.click_input()
                        time.sleep(1)
                    
                    except : 
                        pass
                    time.sleep(1)
                    dlg = app.window(title_re=".*File.*")

                    if(dlg.Button9.get_check_state())==0:
                        dlg.Button9.click_input()
                    dlg.OK.click_input()
                    eng=app.window(title_re=".*Enregistrer.*")
                    eng.Enregistrer.click_input()
                    dlg = app.window(title_re=".*WinOLS.*")
                    dlg.type_keys('^%{ENTER}')
                    dlg = app.window(title_re=".*Project.*")
                    dlg['Edit18'].click_input()
                    dlg['Edit18'].type_keys('Exported')
                    dlg.OK.click_input()
                    dlg = app.window(title_re=".*WinOLS.*")

                while True:
                    try:
                        dlg.Button10.click_input()
                        dlg.MDIClient.children()[0].close()
                        
                    except:
                        break
                count+=1
                time.sleep(10)
if __name__ == "__main__":
        
        keyword= input('keyword de recherche : ')
        path=input("path d'exportation  : ")
        
        bot(keyword,path)