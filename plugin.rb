#SKETCHUP_CONSOLE.clear() #Pour nettoyer la console
require 'csv' #Import les fonctions CSV

  class ExportCSVData

	def initialize()
	end

	def analyze(entity,entityParentName)
	 
		entityLayer = entity.layer 												#Stock le layer
		entityLayerName = entityLayer.name 										#Stock le nom du layer

		if entity.is_a?(Sketchup::ComponentInstance) 							#On récupère le nom de definition
			definitionName = entity.definition.name.split("#").first
		elsif entity.is_a?(Sketchup::Group)
			definitionName = entity.name.split("#").first
		end

		#On check l'entitée et on traite les données si possible
		if entityLayerName != "Layer0" && $visible_layers.include?(entityLayerName) 		#Si le calque n'est pas Layer0, on regarde si le calque est visible

			if entity.is_a?(Sketchup::Group) && entityLayer&.folder&.name == "Cable" 		#si c'est un groupe et que son layer est contenu dans un dossier de layer cable
				cableLength = mesureCable(entity) 											#on mesure le cable
				if cableLength > 0
					definitionName = entityLayerName + " " + cableLength.to_s + "m" 		#on renomme le cable avec sa mesure
					entityLayerName = "Cable"
				else
					definitionName = entityLayerName
					entityLayerName = "Cable Error"
				end
			end

			case $exportType
				when 1 																								#Si $exportType vaut 1, on inclue la localité dans l'export
					ligne = [entityLayerName,definitionName,entity.name.split("#").first,entityParentName] 			#On insère les données dans une ligne
				when 2 																								#Si $exportType vaut 2, on n'inclue pas la localité dans l'export
					ligne = [entityLayerName,definitionName,entity.name.split("#").first] 							#On insère les données dans une ligne
				when 3 																								#Si $exportType vaut 3, on n'inclue pas grand chose
					ligne = [entityLayerName,definitionName] 														#On insère les données dans une ligne
			end

			$definitionsInfos << ligne 																				#On insère la ligne dans un tableau

		end
		 
		#Si il y'a des entitées nestead, on check l'intérieur
		entity.definition.entities.each{|subentity| 				#Pour chaque entity trouvée, on créé un objet subentity
			entityParentName = definitionName
			isGroupOrComponentAnalyse(subentity,entityParentName) 	#On regarde si c'est un groupe ou composant et analyse
		}

	end

	def mesureCable(entity)

		cableLength = 0 									#initialise la taille du cable
		 
		entity.entities.each{ |entity|   					#Pour chaque entity trouvée dans le groupe, on créé un objet Entity

			if !(entity.is_a?(Sketchup::Edge)) 				#s'il y'a un loup dans la bergerie,
				$cableIssues << entity.parent.name 			#on prend note  
				return 0 									#et on abandonnne
			else
				cableLength = cableLength + entity.length 	#On mesure chaque segments
			end
		}

		cableLength = cableLength*0.0254 					#On passe de inch a m

		if cableLength < 5
		cableLength = cableLength.ceil 						#On arrondi au mètre suppérieur les cables de moins de 5m
		elsif cableLength < 30
		cableLength = (cableLength / 5.0).ceil * 5 			#On arrondi les longueurs de moins 30m à multiple de 5 suppérieur
		elsif cableLength > 30
		cableLength = (cableLength / 10.0).ceil * 10 		#On arrondi les longueurs de plus 30m à multiple de 10 suppérieur
		end
		return cableLength 									#On retourne la longeur du cable

	end

	def isVisibleLayerFolder(layerFolder, layerName)
		if layerFolder.visible? 								#Regarde si le dossier de balise est visible
			if layerFolder.folder 								#Regarde si il y'a un dossier de balise dans le dossier de balise
				layerFolder = layerFolder.folder 				#Créé une variable avec le dossier de balise en cours d'analyse
				isVisibleLayerFolder(layerFolder, layerName)	#Appelle la fonction d'analyse de dossier de layer et passe en argument le nom du dossier et de la balise
			else
				$visible_layers << layerName 					#Si il n'y a plus de dossier de balise et que celuici est visible, alors le calque est visible, on insert dans la liste
			end
		end 													#Si le dossier de balise n'est pas visible, on arrête de chercher
	end

	def prepareExportCSV()

		$definitionsInfos = [] 									#On créé une variable pour stocker les infos des definitions (s = array)
		$readyToExport = [] 									#On créé un tableau dans lequel on stockera les infos à exporter
		$cableIssues = [] 										#On créé un tableau des cables qui n'ont pas pu être mesurés

		model = Sketchup.active_model 							#Raccourcis vers le model

		#liste les calques visible
		$visible_layers = [] 									#On créé une variable pour stocker les calques visibles
		model.layers.each { |layer| 							#Pour chacun des layers
			if layer.visible? 									#Si le layer est visible
			layerName = layer.name 								#Créé une variable avec le nom du layer en cours d'analyse
				if layer.folder 								#On regade si il est dans un dossier de layer
				layerFolder = layer.folder 						#Créé une variable avec le dossier de balise en cours d'analyse
				isVisibleLayerFolder(layerFolder, layerName)	# Appelle la fonction d'analyse de dossier de layer et passe en argument le nom du dossier et de la balise
				else
				$visible_layers<< layerName 					# On stock le nom dans la liste des layer visibles
				end
			end
		}

		#On regade chaque entiées
		entities = model.active_entities 						# On récupère toutes les entities dans le model

		entityParentName = ""

		entities.each{ |entity|   								# Pour chaque entity trouvée, on créé un objet Entity
			isGroupOrComponentAnalyse(entity,entityParentName) 	# On regarde si c'est un groupe ou composant et analyse
		}

		entityBilan() 											# Fait le décompte de chaque éléments

		#On génère le CSV

		$readyToExport = $readyToExport.sort_by { |row|  [row[0].to_s, row[1].to_s] } # Error: #<ArgumentError: comparison of Array with Array failed>

		CSV.open("donnees.csv", "wb") do |csv| 											# Création du fichier CSV
		csv << $csvEntete 																# Insertion de l'entête de colonnes
			$readyToExport.each do |ligne| 												# On récupère chaque entrée de definitionsInfos
			csv << ligne 																# On insère chaque ligne dans le fichier csv
			end
		end 																			# Ferme CSV automatiquement
		 
		if $cableIssues
			bilanCalble($cableIssues)
		end

		#on ouvre le fichier csv
		pid = spawn('notepad.exe donnees.csv') 											# Permet l'ouverture du fichier créé
		Process.detach(pid) 															# Détache la commande du programme pour que Sketchup n'attende pas la fermeture du notrepad.

	end

	def isGroupOrComponentAnalyse(subentity,entityParentName) 							# Fonction pour savoir si l'entity est un groupe ou composant, et lancer l'analyse si c'est un groupe ou composant

		if subentity.is_a?(Sketchup::ComponentInstance) || subentity.is_a?(Sketchup::Group) # On cherche à savoir si subentity est un groupe ou un composant
		analyze(subentity,entityParentName) 												# Si oui, on analyze
		end

	end

	def entityBilan()
	 
		qt_sums = Hash.new(0) 																# créer un hash pour stocker les sommes par ligne identique
		 
		$definitionsInfos.each do |ligne| 													# parcourir chaque ligne dans le tableau  
			qt_sums[ligne] += 1 															# ajouter la quantité de l'entité à la somme correspondante dans le hash
		end
		 
		$definitionsInfosRecap = [] 														# créer un nouveau tableau avec les quantités sommées
		$definitionsInfos.each do |ligne|
			ligneRecap = ligne.dup << qt_sums[ligne]   										# ajouter la somme de quantités à la fin de chaque ligne  
			$definitionsInfosRecap << ligneRecap 											# ajouter la ligne modifiée au nouveau tableau
		end

		$readyToExport = $definitionsInfosRecap.uniq 										# Supprime les lignes doublons

	end

	def bilanCalble(liste)

		if liste.length > 0
			@listeHtml = ""
			liste.each{ |cable| @listeHtml = "<li>" + @listeHtml + cable + "</li>"}
			 
			html = '<p><b><u>Export listing cables :</b></u></p>
			<p>Le ou les Groupes suivants n\'ont pas étés interprétés comme des cables :</p><ul>' + @listeHtml +
			'</ul><p>Peut être que ce ou ces groupes ne contiennent pas uniquement des segments</p>
			<button onclick="window.sketchup.close()">Fermer</button>'
		else
			html = '<p><b><u>Export listing cables :</b></u></p>
			<p>Les cables ont bien été listés et exportés</p>
			<button onclick="window.sketchup.close()">Fermer</button>'
		end

		dialog = UI::HtmlDialog.new(
			{
			 :dialog_title => "Info export listing",
			 :preferences_key => "com.sample.plugin",
			 :scrollable => true,
			 :resizable => true,
			 :width => 600,
			 :height => 400,
			 :left => 100,
			 :top => 100,
			 :min_width => 50,
			 :min_height => 50,
			 :max_width =>1000,
			 :max_height => 1000,
			 :style => UI::HtmlDialog::STYLE_DIALOG
			})
		dialog.set_html(html)

		dialog.add_action_callback('close') { |action_context|
		dialog.close
		}

		dialog.show

	end

	def exportBilanLocation()

		$exportType = 1
		$csvEntete = ["Famille","Type","Nom","Localitée","QT"] 						#L'entête du CSV prend en compte la localité
		prepareExportCSV()

	end

	def exportBilanGlobal()

		$exportType = 2
		$csvEntete = ["Famille","Type","Nom","QT"] 									#L'entête du CSV ne prend pas en compte la localité
		prepareExportCSV()

	end

	def exportBilanBrute()

		$exportType = 3
		$csvEntete = ["Famille","Type","QT"] 										#L'entête du CSV Liste brute
		prepareExportCSV()

	end

	def cableTag()

		model = Sketchup.active_model
		layers = model.layers 														# Sélectionne les layers

		layers.each_folder do |folder|
			if folder.name == "Cable"
				result = UI.messagebox("Un dossier nommé 'Cable' a été trouvé.", MB_OK)
				return 																# Si il trouve cable, on arrète tout et on rentre à la maison
			end
		end

		cable_folder = layers.add_folder("Cable") 									# Sinon on créé le dossier "Cable"
		result = UI.messagebox("Un dossier nommé 'Cable' a été créé.", MB_OK)


	end


end

exportdata = ExportCSVData.new

################################
## Creation du menu ############
################################
if( not file_loaded?("exportListCSV.rb") ) 											# Si le fichier n'est pas déjà chargé
	tool_menu = UI.menu("Plugins") 													# on va dans le menu Plugins
	menuCSV = tool_menu.add_submenu("Exports CSV") 									# on ajoute un sous menu
	menuCSV.add_item("Créer Balise Cable") { exportdata.cableTag } 					# on créé un menu Export CSV qui lance la fonction
	menuCSV.add_item("Export Liste Localisée") { exportdata.exportBilanLocation }	# on créé un menu Export CSV qui lance la fonction
	menuCSV.add_item("Export Liste + Nom") { exportdata.exportBilanGlobal } 		# on créé un menu Export CSV qui lance la fonction
	menuCSV.add_item("Export Liste Brute") { exportdata.xexportBilanBrute } 		# on créé un menu Export CSV qui lance la fonction
	file_loaded("exportListCSV.rb") 												# on indique que est chargé
end

