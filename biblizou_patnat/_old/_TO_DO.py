#TODO dans les n2000 hab les codes UE ne correspondent pas au LB_HABDH_FR
#TODO dans les znieff esp les onglets cachés sont mal nommés
#TODO Gestion des erreurs, il arrive qu'il n'y ait pas de ZNIEFF ou pas de N2000 donc pas de xml à télécharger :
# éviter le message d'erreur 2025-02-14 11:32:25,699 - ERROR - Le fichier 'input_xml_n2000_download_list.txt' n'a pas été trouvé dans le dossier \\192.168.1.100\ExEco_Env\Affaires_en_cours\XECO_2403_Carrière Coat Culoden_Rosporden_29\DATA\biblizou_patnat.
#TODO lors du téléchargement des xml : ignorer la 1ere ligne sinon erreur :
# 2025-02-14 11:32:23,838 - ERROR - Erreur HTTP lors du téléchargement de https://inpn.mnhn.fr/docs/ZNIEFF/znieffxml/ID_MNHN.xml: 404 Client Error: Not Found for url: https://inpn.mnhn.fr/docs/ZNIEFF/znieffxml/ID_MNHN.xml
#TODO Le script dans le modeleur ne fonctionne pas en python. Il faut aussi le rendre indépendant de la présence des couches N2000 et ZNIEFF dans le projet.

#TODO dans cette fonction voir les arguments de pivot (surtout aggfunc?)
# La fusion des cellules dans Excel est un comportement qui se produit lors de l'exportation des données,
# et il n'y a pas d'argument dans pivot_table() pour la désactiver directement. Il faut trouver une autre méthode.

#TODO faire les algorithmes de traitement par commune
#TODO essayer d'ajouter la reconstitution des données de la bd status

#TODO znieff et n2000 option à cocher pour tronquer les noms vernaculaires à la virgule.
#TODO renommer les outputs pour plus de lisibilité
#TODO dans le script xml2docx récupérer dans <INCLU_DANS_TYPE2> l'informations d'inclusion des types 1 dans type 2.
# Rajouter un tableau avec :
# - surface
# - Distance et orientation (orientation avec une image, ce serait + fun!)
# - type de zonage (znieff1-znieff2/ZPS-SIC) + zonage parent pour les Znieff1
# - lien vers l'INPN
# - date de création

"HelloWorld - This a place holder for something something I didnt work on yet. Yeayh!"
"Je rajoute des trucs depuis le wouaib, pour voir si je peux le pull sur ma machine. ceci est un test."
"Je rajoute un autre truc depuis mon pc portable à la maison, c'est encore un test !!"

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





          ######                              ###                                           ###
          ##  ##                              ##                                            ##
          ##  ##            ####              ##                ####                        ##      ####
          #####                ##             #####            ##  ##                       ##     ##  ##
          ##  ##            #####             ##  ##           ######                       ##     ######
          ##  ##           ##  ##             ##  ##           ##                           ##     ##
          ######             #####            ######             #####                      ####     #####



                                                             ,---,
                                                           ,--.' |
                                    ,---.                  |  |  :          ,---.           ,---,
                                   '   ,'\                 :  :  :         '   ,'\      ,-+-. /  |
                       ,---.      /   /   |      ,---.     :  |  |,--.    /   /   |    ,--.'|'   |
                      /     \    .   ; ,. :     /     \    |  :  '   |   .   ; ,. :   |   |  ,"' |
                     /    / '    '   | |: :    /    / '    |  |   /' :   '   | |: :   |   | /  | |
                    .    ' /     '   | .; :   .    ' /     '  :  | | |   '   | .; :   |   | |  | |
                    '   ; :__    |   :    |   '   ; :__    |  |  ' | :   |   :    |   |   | |  |/
                    '   | '.'|    \   \  /    '   | '.'|   |  :  :_:,'    \   \  /    |   | |--'
                    |   :    :     `----'     |   :    :   |  | ,'         `----'     |   |/
                     \   \  /                  \   \  /    `--''                      '---'
                      `----'                    `----'


                     ******     ********   *******       ********    ********   *******
                    /*////**   /**/////   /**////**     **//////**  /**/////   /**////**
                    /*   /**   /**        /**   /**    **      //   /**        /**   /**
                    /******    /*******   /*******    /**           /*******   /*******
                    /*//// **  /**////    /**///**    /**    *****  /**////    /**///**
                    /*    /**  /**        /**  //**   //**  ////**  /**        /**  //**
                    /*******   /********  /**   //**   //********   /********  /**   //**
                    ///////    ////////   //     //     ////////    ////////   //     //




                                                                               ......                                                      
                                                                             ..::::.....                                                   
                         ...::::..                  .. ...........         ..::::---::....                                                 
                        .:::::::::::..         ..... .................    ..::::--===:::...                                                
                      ..:---:::::::::::.....................................:::::-=+=--::...                                               
                     .:--===-::::--:::---::................................:::---=+**+=::::...                                             
                    .:--=====:::--:-::---::... ............................::----=++++++=::::::.                                           
                   ..--===+++=::::-:-::-:....   .................:::::.::::::-:--+*++++++=-:::::::.                                        
                  .:--===+++++=::::::-::..       ................::::.:--:::----=*++++++====-:::::::.                                      
                 ..----=++==++=-.::::::..         ..............::::::::::::----=++++==++=---::::::--.                                     
                ..----=++==+===-:.:....           ............::::::-:::::------=+++++=====---::::::::                                     
               .:----=++===--:.::.....            . ..........::..::::::::::----=++++===-----:::::..                                       
              .:---=====--:..  .::...             ...........:..::::::::::::----=++++======-.....                                          
             .:---====--:.      .:...           . .............::.::::::-::-----=++++++===-.                                               
            .:--====--:.       .:-:..  ...       ................:.:::::::::----==++=====-.                                                
            .::.......        ..:-:.. ....       .... ...........:--====-::::---==++===--.                                                 
                               ..::.. ......    . .............:---++====-------=====--:.                                                  
                               ...:....:....     .............:---+=========----====--:..                                                  
                                ..:...:*=::.    .............::---==+++==-=-----===-::....                                                 
                                .... .=%*::. .  .........:..::---=+=#%%#+=----:-+=--:.....                                                 
                               ....  .:++::. . .........:..:-:--===+%%%#+==---:-==--::::..                                                 
                               ...  ...:::..   .......::::.::---===*+++===-::::-==--:::::..                                                
                               ..    .......   .......::::::::-----=+====--:::::---::::::..                                                
                               .    ....   .. ....:..:.:::::------::==-==--::::.::-::::::..                                                
                                     ..   .........:..:::::::--:---::-===-:::::.:::::::::..                                                
                                                ...::....:-------:::::--:::::::.:::::::::..                                                
                                               ..:::.:::-----------:::::::::.::::::::::::..                                                
                                            .....:::::::-:::::::::-:::::::::.:::::::::::::...                                              
                                           .....:::-:::::-::::::::::::::::...::::::::::::.::...                                            
                                           ....::::::------:::::::::::::::...:..:::::::::.::::..                                           
                                           ...::---====----:--::-:::::::::::::..:.::::::::::::.:.                                          
                                  .        ..::---=++===------::-:::::::::::::::.:::::::..:-:::::..                                        
                                  . .    ...:----======++==-----:::.::::::::::::.::::::::.:::::::::..                                      
                            ....   ...    ..:--------===+++=-----:::::::::::::::::::::::.::::::::::::.                                     
                             ....   .... ..:-:--=============-----:.:::::::::::::::::::::::--:::--::::.                                    
                              ....  ......----===++++++++++====--::::::::.:::::::::::::::::-:::----::::.                                   
                          .   ..::.   ...:===++++++++=++++++++=--:::::::.::::::--::--::::::-::-=-:---:::.                                  
                            ...::--:.   .:==+*##+=+*++*###*+++=---::::::::-:::--::----::.::--:==-----:::.                                  
                            ....::--:.. ..==++*+=+++++**#*++++---::::::---:::--::---:::::::-:-=====---::::.                                
                             ....:-==-:...:===+===+++++++++++--:::::::-------:-==--::::::::--=++=--------::.                               
                         .    ...:-=++=-::.:-======++++++=+=-:::::::---:-----=====-::::::-:-=+++===------:::.                              
                         .    ...:--++++=-::.:--====+=====-:::::::--------==+=------::::::-=+**+++++===---::.                              
                         .    ...::-=+**+++--::..:--=----::::::--------===+====--::::::::-=+*#****++=--=--:::.                             
                         .    ....:-=+******+=--:::::::::::---======+++++====----:::::-:-+*#**##**++==-----::.                             
                         .    ....:-=+******+=---==--::::--=++==+++*++++===-=-----:::::-+*#######**+======---:..                           
                         .    ....:--++*+**=.:::----==============++++====---=----:::-:+#########**+++======--::.                          
                         .    ...::--=+*++=:..::::-----==========++++======-----::::--=##########**+++++++===----::::.......               
                         .  ....::--==+*+=:.....:::---=====================-----:::-**###########***++++*+=++============-----::..         
                         . ....:::--==+++:.....::::---============---==++=-----:::-*#######*:  :********+===++++++=============-----:      
                          ...:::---===++-......::::--=====++++==--::-==++=----:::-+#%%###*-     .+*****+=====+++++++++==+++=======----:    
                         ..::::--====++=.......:::--==+++++++=+=-:::-=++++=-:-::=*%%%%%#+.       .*****++++++++***++++++++++=====++=-=++-. 
                        .:::::---==++++:.......:::---====++++++---:::=+++++-:::*%%%%%%#:          -******++++++*+***+++++++++=====+******+:
                        .::---=====+++-.....::::::---===+++++*#--:::-==++++-::-*%%%%#-            .*#*****++++++*+***+++++++++==+==+#%%#+: 
                       .:-::--===++++=...:::--------===++++**#*-------=====-:-++::.                .=*#%##*******++*++++**+*++++=+****+.   
                       ::::---===++++.  ..:::::--=--:=+++*+-:.:---------==--:=-.                        ..:-=++++*+*#********+++******=    
                     ..::::---==++++-           ::.   :.      .--------------:                                       .:--===*###%%%*=:     
                    .:::::--====+++-                          :-------:----::                                                   .          
                   .::::----====++=                          .------=-----:::                                                              
                  .:::::----===++=.                          :---:----:::--:.                                                              
                 .:::-:----=====+:                           :------=:-::---:                                                              
               .:---=----=-======                           .------=+:--:---:                                                              
              .:-=+++-=--=======:                           :---=--*+:-----::                                                              
             .==+*+*+==+========                            --====+#*--====-:                                                              
            :-+*****=-=+++=====:                           .=====+*#*=-====-                                                               
            :++**##*==++**==+=:                            -++===+#%#+=-==+:                                                               
             .-+*#*-******+=:                                .-*####**###+:                                                                
                    .=*#*-                                                                                                                 


''')