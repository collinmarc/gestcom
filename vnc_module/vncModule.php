<?php
require_once '../config/config.inc.php'; // assuming your script is in the root folder of your site

if (!defined('_PS_VERSION_'))
  exit;
 class VncModule extends Module
 {
	 public function __construct()
	  {
		$this->name = 'vncModule';
		$this->tab = 'export';
		$this->version = '1.0';
		$this->author = 'Marc Collin';
		$this->need_instance = 0;
		$this->ps_versions_compliancy = array('min' => '1.6', 'max' => '1.6');
	 
		parent::__construct();
	 
		$this->displayName = 'Vinicom Module';
		$this->description = 'Ce module contient les Addons pour vinicom';
	 
		$this->confirmUninstall = 'Etes-vous sur de vouloir desinstaller ce module?';
	 
	  }
 
	 
	public function install()
		{
		  if (parent::install() == false)
			return false;
		if( !Configuration::updateValue('CSVPATH', ''))
			return false; 
		if( !Configuration::updateValue('REFCMD', 'PRECMD'))
			return false; 
		if( !Configuration::updateValue('ETATCMD', 'PRECOMMANDE'))
			return false; 
		return true;
		}
		
	public function uninstall()
		{
		if (!parent::uninstall()
			|| !Configuration::deleteByName('CSVPATH')
			|| !Configuration::deleteByName('REFCMD')
			|| !Configuration::deleteByName('ETATCMD')
			)
			return false;					
		return true;
		}
	/*********
		Configuration du module
	*********/
	public function getContent(){
		$this->_preProcess();
		$this->_html.='
			<h2>'.$this->displayName.'</h2>
			<form id="vncModule_settings" class="width3" method="post" action="'.$_SERVER['REQUEST_URI'].'">
				<fieldset>					
					<legend><img src="../img/admin/cog.gif" />'.$this->l('Paramètres').'</legend>
					<div class="clear"></div>
					<label>Reference de commandes :</label>
					<div class="margin-form">
						<input name="REFCMD" value="';
							if(Configuration::get('REFCMD') != '')
								$this->_html .= Configuration::get('REFCMD');
							else $this->_html .= '';
						$this->_html .= '">
					</div>
					<div class="clear"></div>
					<label>Etat des commandes :</label>
					<div class="margin-form">
						<input name="ETATCMD" value="';
							if(Configuration::get('ETATCMD') != '')
								$this->_html .= Configuration::get('ETATCMD');
							else $this->_html .= '';
						$this->_html .= '">
					</div>
					<div class="clear"></div>
					<label>Repertoire du fichier CSV</label>
					<div class="margin-form">
						<textarea name="CSVPATH">';
							if(Configuration::get('CSVPATH') != '')
								$this->_html .= Configuration::get('CSVPATH');
							else $this->_html .= '';
						$this->_html .= '</textarea>
					</div>
					<div class="clear"></div>
					<center><input type="submit" name="importPreCommande" value="'.$this->l('Importer PreCommandes').'" class="button" /></center>
				</fieldset>
			</form>
		';
		return $this->_html;
	}
	/*************
		Traitements
	*************/
	private function _preProcess(){		
		$this->_html = "";
		if(Tools::isSubmit('importPreCommande')){				
			if(isset($_POST['CSVPATH']))
			{
				Configuration::updateValue('CSVPATH', $_POST['CSVPATH']);
				Configuration::updateValue('REFCMD', $_POST['REFCMD']);
				Configuration::updateValue('ETATCMD', $_POST['ETATCMD']);
				$result = $this->importPreCommande($_POST['CSVPATH'],$_POST['REFCMD'],$_POST['ETATCMD']);
				$this->_html .= '<div class="conf confirm"><img src="../img/admin/ok.gif" alt="OK" />';
				$this->_html .= 'Les commandes PRECOMMANDES sont  importees';
				$this->_html .= $result;
				$this->_html .= '</div>';
			}
		}
	}//PreProcess


	private $_idshoppingList; // Identifiant du cart
	private $_idCart; // Identifiant du cart
	private $_idOrder; // Identifiant de la commande
	private $_idOrderState; // Identifiant de l'état de la commande
	/*******************************************
	* Import des précommande 
	* Création des ShoppingCart et des commandes correspondant aux précommandes GESTCOM
	********************************************/
	private function importPreCommandesample($pCsvPath)
	{
		$db = MySQLCore::getInstance();
		$db->insert("cart", array(
					'id_customer'=>(int)5,
					'delivery_option'=>'',
					'id_lang'=>'1',
					'id_address_delivery'=>(int)1,
					'id_address_invoice'=>(int)1,
					'id_currency'=>(int)1,
					'id_guest'=>(int)1,
					'date_add'=>date('Y-m-d'),
					'date_upd'=>date('Y-m-d')
					));
		$idCart =$db->Insert_ID();
		
		$product = $this->getProductByReference("563001H");
				$db->insert("cart_product", array(
					'id_cart'=>$idCart,
					'id_product'=>$product->id,
					'date_add'=>date('Y-m-d'),
					'quantity'=>(int)1
					));
					
					
		$product = $this->getProductByReference("214001H");
		$db->insert("cart_product", array(
					'id_cart'=>$idCart,
					'id_product'=>$product->id,
					'date_add'=>date('Y-m-d'),
					'quantity'=>(int)1
					));
		
		$db->insert("orders", array(
					'id_customer'=>(int)5,
					'id_cart'=>$idCart,
					'id_carrier'=>'1',
					'id_lang'=>'1',
					'id_currency'=>(int)1,
					'id_address_delivery'=>(int)1,
					'id_address_invoice'=>(int)1,
					'reference'=>'ORIGINAL4',				
					'current_state'=>'1',
					'payment'=>'',
					'invoice_date'=>date('Y-m-d'),
					'delivery_date'=>date('Y-m-d'),
					'id_shop'=>'1',
					'date_add'=>date('Y-m-d'),
					'date_upd'=>date('Y-m-d')
					));
		$idOrder =$db->Insert_ID();
		
		$product = $this->getProductByReference("563001H");
		$db->insert("order_detail", array(
					'id_order'=>$idOrder,
					'product_id'=>$product->id,
					'id_shop'=>'1',
					'product_name'=>$product->name,
					'product_reference'=>$product->reference,
					'product_weight'=>'0',
					'tax_name'=>'',
					'product_quantity'=>(int)1				
					));
					
		$product = $this->getProductByReference("214001H");
		$db->insert("order_detail", array(
					'id_order'=>$idOrder,
					'product_id'=>$product->id,
					'id_shop'=>'1',
					'product_name'=>$product->name,
					'product_reference'=>$product->reference,
					'product_weight'=>'0',
					'tax_name'=>'',
					'product_quantity'=>(int)1				
					));

	}//importPreCommande

	private function importPreCommande($pCsvPath,$pRefCmd,$pEtatCmd)
	{
		$row = 0;
		$anc_clientId = -1;
		$idCart = -1;
		$result = "";
		$CSVError = str_replace(".csv","ERR.csv",$pCsvPath);
		$fp = fopen($CSVError, 'w');
		if (($handle = fopen($pCsvPath, "r")) !== FALSE) 
		{
			$row = 1;
			$errorline =  array();
			$customer = null;
			$this->_idOrderState = $this->getOrderState($pEtatCmd);
			$result = $result."<br>idOrderSate = ".$this->_idOrderState."<BR>";
			while (($data = fgetcsv($handle, 1000, ";")) !== FALSE) 
			{
				$errorline =  array();
				$errorline[]=$row;
				$result = $result."<br>Traitement LG".$row."<BR>";
				$cltId = (int)$data[0];
				$productref = $data[1];
				$qte = $data[2];
				$pu= $data[3];
				if ($cltId != $anc_clientId)
				{
					$customer = new Customer($cltId);
					if (Validate::isLoadedObject($customer))
					{
						$result = $result."Creation Entete(".$cltId.")<BR>";
						$anc_clientId = $cltId;
						$result = $result.$this->creerEntete($cltId,$pRefCmd);
					}
					else
					{
						$result.=" Client[".$cltId."] inconnu<BR>";
						$errorline[]=" Client[".$cltId."] inconnu";
					}
				}
				if (Validate::isLoadedObject($customer))
				{
					$product = $this->getProductByReference($productref);
					if (Validate::isLoadedObject($product))
					{
						$result = $result."Creation Ligne(".$productref.",".$product->id.",".$product->name.",".$qte.",". $pu.") <BR>";
						$result = $result.$this->creerLigne($productref,$product->id,$product->name, $qte, $pu);
					}
					else
					{
						$result.=" produit[".$productref."] inconnu<BR>";
						$errorline[]=" produit[".$productref."] inconnu";
					}
				}
				if (count($errorline)> 0)
				{
					fputcsv($fp,$errorline);
				}
				
				$row++;

			}//while
			fclose($handle);
			fclose($fp);
		}
		return $result;
	}//importPreCommande
	
	/**
	* Creation des entete de Cart, Order et shoppingList
	**/
	private function creerEntete($pIdCustomer,$pRefCmd)
	{
		$db = MySQLCore::getInstance();
		$this->_idCart = -1;
		$this->_idOrder = -1;
		$db->insert("cart", array(
					'id_customer'=>(int)$pIdCustomer,
					'delivery_option'=>'',
					'id_lang'=>'1',
					'id_address_delivery'=>(int)1,
					'id_address_invoice'=>(int)1,
					'id_currency'=>(int)1,
					'id_guest'=>(int)1,
					'date_add'=>date('Y-m-d'),
					'date_upd'=>date('Y-m-d')
					));
					
		$this->_idCart =$db->Insert_ID();
		$result = "cart[".$this->_idCart."]";
		$test= md5(uniqid(rand(), true));
		$db->insert("orders", array(
					'id_customer'=>(int)$pIdCustomer,
					'id_cart'=>$this->_idCart,
					'id_carrier'=>'1',
					'id_lang'=>'1',
					'id_currency'=>(int)1,
					'id_address_delivery'=>(int)1,
					'id_address_invoice'=>(int)1,
					'secure_key'=>$test,
					'payment'=>"Cheque",
					'module'=>"Cheque",
					'reference'=>$pRefCmd,				
					'current_state'=>(int)$this->_idOrderState,
					'invoice_date'=>date('Y-m-d'),
					'delivery_date'=>date('Y-m-d'),
					'id_shop'=>'1',
					'date_add'=>date('Y-m-d'),
					'date_upd'=>date('Y-m-d')
					));
		$this->_idOrder =$db->Insert_ID();
		// 
		if ($this->_idOrderState >-1)
		{
			$db->insert("order_history", array(
						'id_order'=>$this->_idOrder,
						'id_order_state'=>(int)$this->_idOrderState,
						'id_employee'=>(int)0,
						'date_add'=>date('Y-m-d')
						));
		}
		
		
		
		$db->insert("shopping_list", array(
			'id_customer'=>(int)$pIdCustomer,
			'title'=>$pRefCmd,
			'date_add'=>date('Y-m-d'),
			'date_upd'=>date('Y-m-d'))
			);
		$this->_idshoppingList =$db->Insert_ID();					
		$result = $result."SHL :IdShoppingList(".$this->_idshoppingList.")<BR>";

		
		$result = $result."order[".$this->_idOrder."]<BR>";
		return $result;
	}//createCart
	
	/**
	* Creation des lignes de Cart et Order
	* MAJ des attributs $_idCart $_idOrder
	**/
	private function creerLigne($pproductref, $pProductId,$productname, $pQte, $pPU)
	{
		$result = "";
		$db = MySQLCore::getInstance();
		
		$db->insert("cart_product", array(
					'id_cart'=>$this->_idCart,
					'id_product'=>$pProductId,
					'date_add'=>date('Y-m-d'),
					'quantity'=>$pQte
					));
		
		$db->insert("order_detail", array(
				'id_order'=>$this->_idOrder,
				'product_id'=>$pProductId,
				'id_shop'=>'1',
				'product_name'=>pSQL($productname),
				'product_reference'=>$pproductref,
				'unit_price_tax_incl'=>$pPU,
				'product_weight'=>'0',
				'tax_name'=>'',
				'product_quantity'=>$pQte				
				));
				
		$db->insert("shopping_list_product", array(
			'id_shopping_list'=>(int)$this->_idshoppingList,
			'id_product'=>$pProductId,
			'title'=>pSQL($pproductref."/".$productname)
			));

		return $result;
	}//creerLigne
	
	private function getProductByReference($pReference,Context $context = null)
    {
        if (!$context) {
            $context = Context::getContext();
        }

        $sql = new DbQuery();
        $sql->select('id_product');
        $sql->from('product', 'p');

        $sql->where("p.reference = '".$pReference."'");

        $result = Db::getInstance()->executeS($sql);

        if (!$result) {
            return false;
        }

        $product = null;
        foreach ($result as $row) {
            $product = new Product($row['id_product'],false,1);
        }
        return $product;
    }//getProductByReference
	/**
	* Renvoie l'ID de l'état PRECOMMANDE (ou -1)
	**/
	private function getOrderState($pname)
    {
 
        $sql = new DbQuery();
        $sql->select('id_order_state');
        $sql->from('order_state_lang', 'osl');

        $sql->where("osl.id_lang = 1 AND osl.name = '" . $pname ."'");

        $result = Db::getInstance()->executeS($sql);

        if (!$result) {
            return 0;
        }

        $osl = 0;
        foreach ($result as $row) {
            $osl = $row['id_order_state'];
        }
        return $osl;
    }//getOrderState

}//class vnc_module
?>