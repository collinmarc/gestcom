-- insertion d'un taux de commission pour le client intermédiaire.
insert into FRN_COMM(frnc_FRN_id,FRNC_TYPE_CLIENT,FRNC_TX_COMM)
 SELECT FRN_ID, 'INT', 12.0
FROM FOURNISSEUR where frn_id