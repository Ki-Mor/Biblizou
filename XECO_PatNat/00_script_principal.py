"""
Author : ExEco Environnement
Edition date : 2025/02
Name : 00_script_principal
Group : Biblio_PatNat
"""

print('''
                                                                                                    ..::::.                      
                                                              .------------------=============================================  
                                                         --------------------================================================   
                                                     -------------------=+**###**+==========================================-   
                                                  --=+++++          -==+###########+====         :============+****+========    
                                               +*=--=#+                -############=                 .====+*########*+=====    
                                             =###*=-                      *#*++*##:                      =*############+===     
                                            ######                          =++*=                         ##############+==     
                                          =######                            +#                            +###*==+#####*==     
                                         +++++++                                                            =##====######=      
                                         ==+===                                                              +#*++*######=      
                                       *####*=               ###*                          =###*              ##########+=      
                                      +#######             :+#####                        #######             =########*=       
                                      #######*                   =+                      =+                    ######*+==       
                                     =++****+                     =                      =                     ****++===        
                                     ========                     ==                     =                     =====++++        
                                    =========                     =                      ==                    =+++++++         
                                    ===============================                       =================+++++++++++          
                                   =================++====  ======                         =====  =====+++++++++++++++          
                                   ================+##+==                                          =+++++++++++++++++           
                                  ++====+*+==+++=+++**==                                            -+++***+++++***+            
                                  ##*=+#####+*##*#**##                                                +######++####             
                                  ###+*##+##*=*###=*#+                       :=                        ##*###**##+              
                                 +*===*##+##*=+##+=*.                       -#**.                       ***+++*#+               
                                 ##*==*##+##*==+*==+                       +##**#*                      +###+++                 
                                -=====================                   ++++++++++=                   ++++++                   
                                =========================            =+++++++++++++++++            +++++++                      
                                ==============================+++++++++++++++++++++++++++++++++++++++++=                        
                               ===========================++++++++++++++++++++++++++++++++++++++++++                            
                              =======================++++++++++++++++++++++++++++++++++++++++++                                 
                              ===================--::..::::::------======+++++++=.                                              
     
                     ##   ##                              ###
                     ##   ##                               ##
                     ##   ##            ####               ##               ####              ####             ##  ##             ####
                     ## # ##           ##  ##              ##              ##  ##            ##  ##            #######           ##  ##
                     #######           ######              ##              ##                ##  ##            ## # ##           ######
                     ### ###           ##                  ##              ##  ##            ##  ##            ##   ##           ##
                     ##   ##            #####             ####              ####              ####             ##   ##            #####
            
            
                                                                        ## ##
                                                                        ##  ##
                                                                            ##
                                                                           ##
                                                                          ##
                                                                         #   ##
                                                                        ######


 ##  ##            #######             ####             #####                     ######                                                 ###
 ##  ##             ##   #            ##  ##           ##   ##                    # ## #                                                  ##
  ####              ## #             ##                ##   ##                      ##               ####              ####               ##               #####
   ##               ####             ##                ##   ##                      ##              ##  ##            ##  ##              ##              ##
  ####              ## #             ##                ##   ##                      ##              ##  ##            ##  ##              ##               #####
 ##  ##             ##   #            ##  ##           ##   ##                      ##              ##  ##            ##  ##              ##                   ##
 ##  ##            #######             ####             #####                      ####              ####              ####              ####             ######


                         ######                                ##              ##   ##                               ##
                          ##  ##                               ##              ###  ##                               ##
                          ##  ##            ####              #####            #### ##            ####              #####
                          #####                ##              ##              ## ####               ##              ##
                          ##                #####              ##              ##  ###            #####              ##
                          ##               ##  ##              ## ##           ##   ##           ##  ##              ## ##
                         ####               #####               ###            ##   ##            #####               ###


                                                                                                                    
''')

import subprocess
import sys
import time

# Définir une fois le chemin du dossier
folder_path = input("Entrez le chemin du dossier de travail : ")

# Temps global de début
start_time_global = time.time()

# Liste des scripts à exécuter
scripts = [
    # "01_package_verification.py",
    # "03_inputs_xlsx2txt.py",
    "11_znieff_xml_download_list.py",
    "12_znieff_xml2docx_com.py",
    "13_znieff_xml2xlsx_esp.py",
    "14_znieff_xml2xlsx_hab.py"
    "21_n2000_xml_download_list.py",
    "22_n2000_xml2docx_desc.py",
    "23_n2000_xml2xlsx_esp.py",
    "24_n2000_xml2xlsx_hab.py"
    # "99_Delete_useless_files.py"
]

# Fonction pour exécuter un script en passant le chemin du dossier
def execute_script(script_name, folder_path):
    # Temps de début pour chaque script
    start_time_script = time.time()

    try:
        # Exécuter le script en passant le chemin comme argument
        subprocess.check_call([sys.executable, script_name, folder_path])
        print(f"{script_name} a été exécuté avec succès.")
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de l'exécution de {script_name}: {e}")
    finally:
        # Temps de fin pour chaque script et calcul du temps d'exécution
        end_time_script = time.time()
        elapsed_time_script = end_time_script - start_time_script
        print(f"Temps d'exécution de {script_name}: {elapsed_time_script:.2f} secondes")

# Boucle pour exécuter chaque script avec le chemin du dossier
for script in scripts:
    execute_script(script, folder_path)

# Temps de fin global
end_time_global = time.time()
elapsed_time_global = end_time_global - start_time_global
print(f"\nTemps total d'exécution de tous les scripts : {elapsed_time_global:.2f} secondes")
