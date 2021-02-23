from pywinauto.application import Application
from PIL import Image
from PIL import ExifTags
import os
from time import sleep
import shutil
import pathlib

def takeSeq(elem):
    return elem[0]

def get_tags(tags,dir):
    for entry in os.listdir(dir):
        if(entry.lower().endswith('.jpg')):
            my_pic = Image.open(dir+'\\'+entry)
            if(my_pic._getexif() != None):
                exif_dic = { ExifTags.TAGS[k]: v for k, v in my_pic._getexif().items() if k in ExifTags.TAGS }
                if('XPKeywords' in exif_dic.keys()):
                    tags.append([entry,exif_dic['XPKeywords'].decode('utf-8').replace('\0','')])
                else:
                    tags.append([entry,str(None)])
            else:
                tags.append([entry,str(None)])
    tags.sort(key=takeSeq)

def write_to_txt(tags,fd):
    for item in tags:
        fd.write(item[0]+'->'+str(item[1])+'\n')

conf_ini = open('config.ini','r')

conf_list = conf_ini.readlines()

APT_exe_path = conf_list[0].replace('\n','')
raw_pic_dir = conf_list[1].replace('\n','')
Sensitivity_conf = conf_list[2].replace('\n','').lower().capitalize()
Ignore_small_pic = conf_list[3].replace('\n','').lower()

if(Ignore_small_pic == 'true'):Ignore_small_pic =True
else : Ignore_small_pic = False

raw_tmp_dir = os.path.join(os.path.expandvars("%userprofile%"),"Desktop\\__pic_with_correct_tags__")
APT_tmp_dir = os.path.join(os.path.expandvars("%userprofile%"),"Desktop\\__pic_without_tags__")

pathlib.Path(raw_tmp_dir).mkdir(parents=True, exist_ok=True)
pathlib.Path(APT_tmp_dir).mkdir(parents=True, exist_ok=True)

print("Preprocessing Materials...\nPlease Wait......")

shutil.copytree(raw_pic_dir,raw_tmp_dir, dirs_exist_ok=True)

for entry in os.listdir(raw_tmp_dir):
    if entry.lower().endswith('.jpg'):
        image = Image.open(raw_tmp_dir+'\\'+entry)
        data = list(image.getdata())
        image_without_exif = Image.new(image.mode, image.size)
        image_without_exif.putdata(data)
        image_without_exif.save(APT_tmp_dir+'\\'+entry)

if(Ignore_small_pic):
    Threshold = conf_list[4].replace('\n','')
    for entry in os.scandir(APT_tmp_dir):
        if(os.path.getsize(entry.path)<float(Threshold)*1024):
            os.remove(entry.path)
            os.remove(entry.path.replace('out','_correct'))

print("Done!\n")

raw_tags_list = []
get_tags(raw_tags_list,raw_tmp_dir)

print("\nAPT is running...\nPlease Wait......\n")

app = Application(backend = 'uia').start(APT_exe_path)
app.dialog.Options.click()

app.dialog.dialog2.sensitivitycombobox.select(Sensitivity_conf)

Ignore_Status = app.dialog.dialog2.checkbox0.get_toggle_state()
if(Ignore_Status != Ignore_small_pic):
    app.dialog.dialog2.checkbox0.toggle()
    Ignore_status = Ignore_small_pic
if(Ignore_small_pic):
    app.dialog.dialog2.edit.set_text(Threshold)
app.dialog.dialog2.ok.click()

app.dialog.add.click()
app.dialog.BrowseForFolder.Pane.TreeView.get_item(r'\Desktop\pic_without_tags').select()
app.dialog.BrowseForFolder.OK.click()
app.dialog.Analyze.click()

app.dialog.dialog2.wait('enabled',timeout=10000000)

app.dialog.dialog2.OK.click()
app.dialog.SaveTags.click()
app.dialog.dialog2.yes.click()

print("\nSaving tags\n")

app.dialog.SaveTags.wait('exists',timeout=10000000)

app.dialog.close()

APT_tags_list=[]
get_tags(APT_tags_list,APT_tmp_dir)

print("\nAnalysis finished!\n")

print('\nOutputing result...\n')

count = 0
#res_txt = open(os.path.join(os.path.expandvars("%userprofile%"),"Desktop\\res.txt"),'w')
res_txt = open(('res.txt'),'w')

for x,y in zip(raw_tags_list,APT_tags_list):
    if(y[1] not in x[1]):
        res_txt.write('Failed case {}:\n'.format(count+1)+x[0]+'\n'+'Raw:'+' -> '+str(x[1])+'\n'+'APT:'+' -> '+str(y[1])+'\n\n')
        count=count+1

acc = (len(raw_tags_list)-count)/len(raw_tags_list)

res_txt.write('Total Cases Count = {}\n'.format(len(raw_tags_list)))
res_txt.write('Total Failed Cases Count = {}\n'.format(count))
res_txt.write('Accuracy Rate = '+str(acc*100)+r'%'+'\n')

shutil.rmtree(raw_tmp_dir)
shutil.rmtree(APT_tmp_dir)

print('\nResult has been successfully saved to '+r'{}'.format(os.getcwd()+'\\res.txt')+'\n')