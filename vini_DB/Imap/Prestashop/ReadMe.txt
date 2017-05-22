Ces fichiers sont les fichiers nécessaires à prestashop pour envoyer les commandes à Gestcom
new_order.txt => [documentRoot]\themes\default-bootstrap\modules\mailalerts\mails\fr
mailalerts.php =>[documentRoot]\modules\mailalerts

Paramétrage : 
Paramètres avancés / Email 
* Utiliser mes propres paramètres SMTP
* Envoyer les deux Formats de messages (HTML et texte)

Vérification
Modules / MailAlerts : 
	* New Order = OUI
		E-mail addresses : mail d'import de commandes prestashop

Localisation / Traductions
	Type de traduction = "Traduction des modèles d'emails"
	Selectionez votre theme = "default-bootstrap"
	langue  = "fr"

	mailalerts / new_order / Version texte
[?xml version="1.0" encoding="utf-8" standalone="yes"?]
[cmdprestashop]
[id]{order_id}[/id]
[name]{order_name}[/name]
[customer_id]{customer_id}[/customer_id]
		[livraison_company]{delivery_company}[/livraison_company]
		[livraison_name]{delivery_company}[/livraison_name]
		[livraison_firstname]{delivery_company}[/livraison_firstname]
		[livraison_adress1]{delivery_address1}[/livraison_adress1]
		[livraison_adress2]{delivery_address2}[/livraison_adress2]
		[livraison_postalcode]{delivery_postal_code}[/livraison_postalcode]
		[livraison_city]{delivery_city}[/livraison_city]
[lignes]
{itemsVnc}
[/lignes]
[/cmdprestashop]
[/xml]


