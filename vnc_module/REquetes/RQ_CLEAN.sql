DELETE FROM pre8466_cart_product where id_cart in (SELECT id_cart FROM `pre8466_orders` WHERE id_order >123);
DELETE FROM pre8466_cart where id_cart in (SELECT id_cart FROM `pre8466_orders` WHERE id_order >123);
DELETE FROM pre8466_order_detail  WHERE id_order >123;
DELETE FROM pre8466_orders  WHERE id_order >123;
DELETE FROM pre8466_shopping_list_product;
DELETE FROM pre8466_shopping_list;
