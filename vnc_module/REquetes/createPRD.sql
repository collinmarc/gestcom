--
-- Structure de la table `pre8466_product`
--

CREATE TABLE `pre8466_product2` (
  `id_product` int(10) UNSIGNED NOT NULL,
  `id_supplier` int(10) UNSIGNED DEFAULT NULL,
  `id_manufacturer` int(10) UNSIGNED DEFAULT NULL,
  `id_category_default` int(10) UNSIGNED DEFAULT NULL,
  `id_shop_default` int(10) UNSIGNED NOT NULL DEFAULT '1',
  `id_tax_rules_group` int(11) UNSIGNED NOT NULL,
  `on_sale` tinyint(1) UNSIGNED NOT NULL DEFAULT '0',
  `online_only` tinyint(1) UNSIGNED NOT NULL DEFAULT '0',
  `ean13` varchar(13) DEFAULT NULL,
  `upc` varchar(12) DEFAULT NULL,
  `ecotax` decimal(17,6) NOT NULL DEFAULT '0.000000',
  `quantity` int(10) NOT NULL DEFAULT '0',
  `minimal_quantity` int(10) UNSIGNED NOT NULL DEFAULT '1',
  `price` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `wholesale_price` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `unity` varchar(255) DEFAULT NULL,
  `unit_price_ratio` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `additional_shipping_cost` decimal(20,2) NOT NULL DEFAULT '0.00',
  `reference` varchar(32) DEFAULT NULL,
  `supplier_reference` varchar(32) DEFAULT NULL,
  `location` varchar(64) DEFAULT NULL,
  `width` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `height` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `depth` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `weight` decimal(20,6) NOT NULL DEFAULT '0.000000',
  `out_of_stock` int(10) UNSIGNED NOT NULL DEFAULT '2',
  `quantity_discount` tinyint(1) DEFAULT '0',
  `customizable` tinyint(2) NOT NULL DEFAULT '0',
  `uploadable_files` tinyint(4) NOT NULL DEFAULT '0',
  `text_fields` tinyint(4) NOT NULL DEFAULT '0',
  `active` tinyint(1) UNSIGNED NOT NULL DEFAULT '0',
  `redirect_type` enum('','404','301','302') NOT NULL DEFAULT '',
  `id_product_redirected` int(10) UNSIGNED NOT NULL DEFAULT '0',
  `available_for_order` tinyint(1) NOT NULL DEFAULT '1',
  `available_date` date NOT NULL DEFAULT '0000-00-00',
  `condition` enum('new','used','refurbished') NOT NULL DEFAULT 'new',
  `show_price` tinyint(1) NOT NULL DEFAULT '1',
  `indexed` tinyint(1) NOT NULL DEFAULT '0',
  `visibility` enum('both','catalog','search','none') NOT NULL DEFAULT 'both',
  `cache_is_pack` tinyint(1) NOT NULL DEFAULT '0',
  `cache_has_attachments` tinyint(1) NOT NULL DEFAULT '0',
  `is_virtual` tinyint(1) NOT NULL DEFAULT '0',
  `cache_default_attribute` int(10) UNSIGNED DEFAULT NULL,
  `date_add` datetime NOT NULL,
  `date_upd` datetime NOT NULL,
  `advanced_stock_management` tinyint(1) NOT NULL DEFAULT '0',
  `pack_stock_type` int(11) UNSIGNED NOT NULL DEFAULT '3'
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- update structure
ALTER TABLE `pre8466_product2` ADD PRIMARY KEY(`id_product`);
ALTER TABLE `pre8466_product2` CHANGE `id_product` `id_product` INT(10) UNSIGNED NOT NULL AUTO_INCREMENT;
ALTER TABLE `pre8466_image_shop` DROP PRIMARY KEY, ADD PRIMARY KEY (`id_image`, `id_shop`, `id_product`) USING BTREE;
ALTER TABLE `pre8466_image` DROP PRIMARY KEY, ADD PRIMARY KEY (`id_image`, `id_product`) USING BTREE;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `getidpafromean13`  AS  select min(`pa2`.`id_product_attribute`) AS `id_product_attribute`,`pa2`.`ean13` AS `ean13` from `pre8466_product_attribute` `pa2` where (`pa2`.`location` = '') group by `pa2`.`ean13` order by `pa2`.`ean13` ;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `getidproduct2fromean13`  AS  select min(`pre8466_product2`.`id_product`) AS `id_product`,`pre8466_product2`.`ean13` AS `ean13` from `pre8466_product2` where (`pre8466_product2`.`location` = '') group by `pre8466_product2`.`ean13` ;

--creation de product2
UPDATE `pre8466_product` SET ean13 = id_product;
DELETE from pre8466_product2;
INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>1);


INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>2);

INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>3);

INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>4);

INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>5);

INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>6);

INSERT INTO
  `pre8466_product2`(
    `id_supplier`,
    `id_manufacturer`,
    `id_category_default`,
    `id_shop_default`,
    `id_tax_rules_group`,
    `on_sale`,
    `online_only`,
    `ean13`,
    `upc`,
    `ecotax`,
    `quantity`,
    `minimal_quantity`,
    `price`,
    `wholesale_price`,
    `unity`,
    `unit_price_ratio`,
    `additional_shipping_cost`,
    `reference`,
    `supplier_reference`,
    `location`,
    `width`,
    `height`,
    `depth`,
    `weight`,
    `out_of_stock`,
    `quantity_discount`,
    `customizable`,
    `uploadable_files`,
    `text_fields`,
    `active`,
    `redirect_type`,
    `id_product_redirected`,
    `available_for_order`,
    `available_date`,
    `condition`,
    `show_price`,
    `indexed`,
    `visibility`,
    `cache_is_pack`,
    `cache_has_attachments`,
    `is_virtual`,
    `cache_default_attribute`,
    `date_add`,
    `date_upd`,
    `advanced_stock_management`,
    `pack_stock_type`
  )
SELECT
  `id_supplier`,
  `id_manufacturer`,
  `id_category_default`,
  `id_shop_default`,
  `id_tax_rules_group`,
  `on_sale`,
  `online_only`,
  `ean13`,
  `upc`,
  `ecotax`,
  `quantity`,
  `minimal_quantity`,
  `price`,
  `wholesale_price`,
  `unity`,
  `unit_price_ratio`,
  `additional_shipping_cost`,
  `reference`,
  `supplier_reference`,
  `location`,
  `width`,
  `height`,
  `depth`,
  `weight`,
  `out_of_stock`,
  `quantity_discount`,
  `customizable`,
  `uploadable_files`,
  `text_fields`,
  `active`,
  `redirect_type`,
  `id_product_redirected`,
  `available_for_order`,
  `available_date`,
  `condition`,
  `show_price`,
  `indexed`,
  `visibility`,
  `cache_is_pack`,
  `cache_has_attachments`,
  `is_virtual`,
  `cache_default_attribute`,
  `date_add`,
  `date_upd`,
  `advanced_stock_management`,
  `pack_stock_type`
FROM
  pre8466_product
  WHERE id_product in(SELECT id_product  FROM `pre8466_product_attribute` 
group by id_product
having count(*)>7);


--- MAJ de l'id du product (pour éviter les conflits avec product)
UPDATE
  `pre8466_product2`
SET
  id_product = id_product +(
  SELECT
    MAX(id_product)+10
  FROM
    pre8466_product
) ; 


  
-- MAJ du ean13 dans product (ancienid)
UPDATE `pre8466_product_attribute` SET `ean13`= "";
UPDATE `pre8466_product_attribute` SET `ean13` = id_product WHERE default_on IS NULL AND id_product IN( SELECT DISTINCT ean13 FROM pre8466_product2 );



-- MAJ de product_attrivute
 UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

 
--
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

 UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

-- 
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

 --
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

 UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

-- 
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

 UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

--
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

 UPDATE
  pre8466_product_attribute pa
SET
  location = "X",
  upc =(
select id_product from getidproduct2fromean13 where ean13 = pa.ean13)
WHERE
  pa.ean13 <> "" AND pa.location = "" AND pa.id_product_attribute =(
  SELECT id_product_attribute
  FROM
    getidpafromean13 v1
  WHERE
    v1.ean13 = pa.ean13
);

-- MAJ de la reference
update pre8466_product2 p2 set location = "X" where id_product in (
select id_product from getidproduct2fromean13 where ean13 = p2.ean13);

update `pre8466_product2` p2
set reference = (select reference from pre8466_product_attribute pa where pa.upc = p2.id_product);

UPDATE `pre8466_product_attribute` SET `id_product` = `upc`, location = "Y" where upc <> "";

UPDATE `pre8466_product_attribute` SET `default_on` = 1 where location = "Y";

-- recopie des infos dans product
INSERT INTO `pre8466_product`(`id_product`, `id_supplier`, `id_manufacturer`, `id_category_default`, `id_shop_default`, `id_tax_rules_group`, `on_sale`, `online_only`, `ean13`, `upc`, `ecotax`, `quantity`, `minimal_quantity`, `price`, `wholesale_price`, `unity`, `unit_price_ratio`, `additional_shipping_cost`, `reference`, `supplier_reference`, `location`, `width`, `height`, `depth`, `weight`, `out_of_stock`, `quantity_discount`, `customizable`, `uploadable_files`, `text_fields`, `active`, `redirect_type`, `id_product_redirected`, `available_for_order`, `available_date`, `condition`, `show_price`, `indexed`, `visibility`, `cache_is_pack`, `cache_has_attachments`, `is_virtual`, `cache_default_attribute`, `date_add`, `date_upd`, `advanced_stock_management`, `pack_stock_type`) 
SELECT `id_product`, `id_supplier`, `id_manufacturer`, `id_category_default`, `id_shop_default`, `id_tax_rules_group`, `on_sale`, `online_only`, `ean13`, `upc`, `ecotax`, `quantity`, `minimal_quantity`, `price`, `wholesale_price`, `unity`, `unit_price_ratio`, `additional_shipping_cost`, `reference`, `supplier_reference`, `location`, `width`, `height`, `depth`, `weight`, `out_of_stock`, `quantity_discount`, `customizable`, `uploadable_files`, `text_fields`, `active`, `redirect_type`, `id_product_redirected`, `available_for_order`, `available_date`, `condition`, `show_price`, `indexed`, `visibility`, `cache_is_pack`, `cache_has_attachments`, `is_virtual`, `cache_default_attribute`, `date_add`, `date_upd`, `advanced_stock_management`, `pack_stock_type` from pre8466_product2;

INSERT INTO `pre8466_product_lang`(`id_product`, `id_shop`, `id_lang`,  `name`,`link_rewrite`)
SELECT p2.id_product,1,1 , pl.name, pl.`link_rewrite`
from pre8466_product2 p2,pre8466_product_lang pl
where p2.ean13 = pl.id_product
;

INSERT INTO `pre8466_product_shop`(`id_product`, `id_shop`, `id_category_default`,id_tax_rules_group ,date_add,date_upd,active,price)
SELECT p2.id_product,1 , ps.id_category_default, ps.id_tax_rules_group, ps.date_add,ps.date_upd,ps.active, ps.price
from pre8466_product2 p2,pre8466_product_shop ps
where p2.ean13 = ps.id_product
;

update pre8466_product_shop
SET active = 1 where id_product in (select id_product from pre8466_product2);



-- les nouveau produit sont dans la même catégorie
 
INSERT INTO `pre8466_category_product`(id_product,`id_category`,position )
SELECT p2.id_product, cp.id_category, cp.position
from pre8466_product2 p2,pre8466_category_product cp
where p2.ean13 = cp.id_product;


update pre8466_product_attribute_shop pas
SET 
pas.id_product = (select id_product from pre8466_product_attribute pa where pa.id_product_attribute = pas.id_product_attribute);

UPDATE pre8466_product_attribute_shop pas SET default_on = 1 WHERE id_product IN( SELECT id_product FROM pre8466_product2 );

-- Stock
update pre8466_stock_available sa
SET 
sa.id_product = (select id_product from pre8466_product_attribute pa where pa.id_product_attribute = sa.id_product_attribute)
where id_product in ( select distinct ean13 from pre8466_product2 )  and id_product_attribute <>0;

INSERT INTO pre8466_stock_available( `id_product`, `id_product_attribute`, `id_shop`, `id_shop_group`, `quantity`, `depends_on_stock`, `out_of_stock` ) SELECT `id_product`, 0, `id_shop`, `id_shop_group`, `quantity`, `depends_on_stock`, `out_of_stock` FROM `pre8466_stock_available` WHERE id_product IN( SELECT id_product FROM pre8466_product2 );

-- image 
INSERT INTO `pre8466_image`(`id_image`,`id_product`,position,cover ) SELECT pis.id_image,p2.id_product, pis.position,pis.cover from 
pre8466_product2 p2,pre8466_image pis where p2.ean13 = pis.id_product ;

INSERT INTO `pre8466_image_shop`( `id_image`, `id_product`, id_shop, cover ) SELECT pis.id_image, p2.id_product, 1, pis.cover FROM pre8466_product2 p2, pre8466_image_shop pis WHERE p2.ean13 = pis.id_product ;

-- Description

update pre8466_product_lang pl
SET description_short = concat(description_short,"  ",(
SELECT
  al.name
FROM
  pre8466_product2 p2,
  pre8466_product_attribute pa,
  pre8466_product_attribute_combination pac,
  pre8466_attribute_lang al,
  pre8466_attribute a
WHERE
 pa.id_product = p2.id_product AND pa.default_on = 1 AND
 pac.id_product_attribute = pa.id_product_attribute AND 
 al.id_attribute = pac.id_attribute AND a.id_attribute = pac.id_attribute AND a.id_attribute_group = 3
    And p2.id_product = pl.id_product)
 where pl.id_product in (select id_product from pre8466_product2);
 
 update pre8466_product_lang pl
SET description_short = concat(description_short," / ",(
SELECT
  al.name
FROM
  pre8466_product2 p2,
  pre8466_product_attribute pa,
  pre8466_product_attribute_combination pac,
  pre8466_attribute_lang al,
  pre8466_attribute a
WHERE
 pa.id_product = p2.id_product AND pa.default_on = 1 AND
 pac.id_product_attribute = pa.id_product_attribute AND 
 al.id_attribute = pac.id_attribute AND a.id_attribute = pac.id_attribute AND a.id_attribute_group = 1
    And p2.id_product = pl.id_product))
 where pl.id_product in (select id_product from pre8466_product2);
 ---UPDATE 22/12/2016
 -- RAZ des descriptions
 DELETE from pre8466_product_lang
where id_product in (select id_product from pre8466_product2);
--Recopie des descritions des produits originaux
INSERT
INTO
  `pre8466_product_lang`(
    `id_product`,
    `id_shop`,
    `id_lang`,
    `name`,
    `link_rewrite`,
    description,
    description_short,
    meta_description,
    meta_keywords,
    meta_title,
    available_now,
    available_later
  )
SELECT
  p2.id_product,
  1,
  1,
  pl.name,
  pl.`link_rewrite`,
  pl.description,
  pl.description_short,
  pl.meta_description,
  pl.meta_keywords,
  pl.meta_title,
  pl.available_now,
  pl.available_later
FROM
  pre8466_product2 p2,
  pre8466_product_lang pl
WHERE
  p2.ean13 = pl.id_product
  
-- Description

UPDATE
  pre8466_product_lang pl
SET
  description_short = CONCAT(
    description_short,
    "  ",
    (
    SELECT
      al.name
    FROM
      pre8466_product2 p2,
      pre8466_product_attribute pa,
      pre8466_product_attribute_combination pac,
      pre8466_attribute_lang al,
      pre8466_attribute a
    WHERE
      pa.id_product = p2.id_product AND pa.default_on = 1 AND pac.id_product_attribute = pa.id_product_attribute AND al.id_attribute = pac.id_attribute AND a.id_attribute = pac.id_attribute AND a.id_attribute_group = 3 AND p2.id_product = pl.id_product
  ))
WHERE
  pl.id_product IN(
  SELECT
    id_product
  FROM
    pre8466_product2
)
 
 update pre8466_product_lang pl
SET description_short = concat(description_short," / ",(
SELECT
  al.name
FROM
  pre8466_product2 p2,
  pre8466_product_attribute pa,
  pre8466_product_attribute_combination pac,
  pre8466_attribute_lang al,
  pre8466_attribute a
WHERE
 pa.id_product = p2.id_product AND pa.default_on = 1 AND
 pac.id_product_attribute = pa.id_product_attribute AND 
 al.id_attribute = pac.id_attribute AND a.id_attribute = pac.id_attribute AND a.id_attribute_group = 1
    And p2.id_product = pl.id_product))
 where pl.id_product in (select id_product from pre8466_product2);

 -- feature
 INSERT
INTO
  `pre8466_feature_product`(
    `id_product`,
    `id_feature`,
    `id_feature_value`
  )
SELECT
  p2.id_product,
  pf.`id_feature`,
  pf.`id_feature_value`
FROM
  pre8466_product2 p2,
  pre8466_feature_product pf
WHERE
  p2.ean13 = pf.id_product
;
 -- attachment
INSERT
INTO
  `pre8466_product_attachment`(
    `id_product`,
    `id_attachment`
  )
SELECT
  p2.id_product,
  pa.`id_attachment`
FROM
  pre8466_product2 p2,
  pre8466_product_attachment pa
WHERE
  p2.ean13 = pa.id_product;
  