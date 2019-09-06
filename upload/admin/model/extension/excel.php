<?php class ModelExtensionExcel extends Model {
    public $defaultLang;
    public $stores = [];
    public function import($filename, $row_order = 0) {
        $this->load->language('catalog/excel');
        $filename = DIR_UPLOAD.$filename;
        require_once DIR_SYSTEM.'library/PHPExcel/Classes/PHPExcel.php';

        $cacheMethod    = PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized;
		$inputFileType  = PHPExcel_IOFactory::identify($filename);
        $objReader      = PHPExcel_IOFactory::createReader($inputFileType);
        
		$objReader->setReadDataOnly(true);
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
        
		$reader    = $objReader->load($filename);
        $import    = $reader->getSheet(0)->toArray(null, true, true, true);
   
        $this->load->model('setting/setting');
        $config = $this->model_setting_setting->getSetting('eimport');
        
        $this->load->model('setting/store');
        $this->load->model('catalog/category');

		$this->stores[] = '0';
        
        $stores_all = $this->model_setting_store->getStores();

		foreach ($stores_all as $store) {
			$this->stores[] = $store['store_id'];
		}


        $status = [
            'total' => 0, 'page' => 0, 'logs' => []
        ];

        /** get products option rows in some array */
        $import_data = [];
        $imported_data = [];
        $firstRow = isset($config['eimport_firstRow']) && 1 == $config['eimport_firstRow'] ? 1 : 0;
        forEach($import as $ik => $iv) {
            $allow = $ik == 1 && $firstRow == 1 ? false : true;

            if($allow) {
                $model = $iv[$config['eimport_model']];
                if( !in_array($model, $imported_data) && "" != trim($model) ) {
                    $iv['options'] = [];
                    $import_data[$model] = $iv;
                    $imported_data[] = $model;
                }elseif( "" != trim($model) ) {
                    $import_data[$model]['options'][] = $iv;
                }else{
                    $this->session->data['error'][] = $ik.$this->language->get('ierr_no_model');
                }
            }
            
        }
        /** get products option rows in some array */



        $first_row = true;
        
        $start     = $row_order;
        $stop      = count($import_data);
        $import_keys = array_keys($import_data);
        $current   = 0;
        $limit     = 50;
        
        $status['total'] = $stop;


        for($i = $start; $i<$stop; $i++) {
            $model = $import_keys[$i];
            $product_images = $this->getProductImages($model);
            $product_image = isset($product_images[0]) ? $product_images[0]['image'] : '';
            if(count($product_images) > 1) array_splice($product_images, 0, 1);


            /** product description */
            $desc_metod= $config['eimport_ddMetod'];
            $desc_text = $config['eimport_defaultDesc'];
            $desc_row  = $config['eimport_description'];

            $description = '';

            if($desc_metod == 1) $description = $import_data[$model][$desc_row];

            if($desc_metod == 2) $description = $desc_text;

            if($desc_metod == 3 ) {

                if("" == trim($import_data[$model][$desc_row]) ) $description = $desc_text;
                else $description = $import_data[$model][$desc_row];

            }

            $re = '/{([A-Z]+)}/m';
            preg_match_all($re, $description, $matches, PREG_SET_ORDER, 0);
            
            if(!empty($matches)) {
                forEach($matches as $match) {
                    $target_row = isset($match[1]) ? $match[1] : '';

                    if( $target_row != '' && isset($import_data[$model][$target_row]) ) { $description = str_replace("{".$match[1]."}", $import_data[$model][$target_row], $description ); }
                }
            }

            /** product name */
            $name        = '';
            $name_row    = $config['eimport_name'];
            $name_source = $import_data[$model][$name_row];
            
            $name_prefix = $config['eimport_productNamePrefixMetod'];
            $prefix_row  = $config['eimport_productNamePrefixSource'];
            $name_sufix  = $config['eimport_productNameSufixMetod'];
            $suffix_row  = $config['eimport_productNameSufixSource'];


            if($name_prefix == 0 && "" == trim($name_source)) $name.= $import_data[$model][$prefix_row];
            if($name_prefix == 1) $name.= $import_data[$model][$prefix_row];

            if("" != trim($name_source)) $name.= ' '.$name_source;

            if($name_sufix == 0 && "" == trim($name_source)) $name.= $import_data[$model][$prefix_row];
            if($name_sufix == 1) $name.= $import_data[$model][$prefix_row];


            $meta_desc  = '';
            $meta_metod = $config['eimport_metaDescMetod'];
            $meta_row   = $config['eimport_metaDesc'];

            if( $meta_metod == '0' ) {
                $meta_desc = $import_data[$model][$meta_row];
            }

            if( $meta_metod == '1' ) {
                $meta_desc = explode('\n', $import_data[$model][$meta_row]);
                $meta_desc = $meta_desc[0];
            }

            if( $meta_metod == '2' ) {
                $meta_desc = mb_substr(strip_tags($description), 0, 150);
            }


            $tags = isset($config['eimport_metaTags']) && "00" != $config['eimport_metaTags'] ? $import_data[$model][$config['eimport_metaTags']] : '';
            $this->defaultLang = (int)$this->config->get('config_language_id');

            $product_description[$this->defaultLang] = [
                'name' => $name,
                'description' => $description,
                'meta_description' => $meta_desc,
                'tag' => $tags,
                'meta_title' => $name,
                'meta_keyword' => ''
            ];

            /** product_attributes */
            $product_attributes = [];
            $attrs = $config['eimport_attrs'];

            forEach($attrs as $attr_id => $attr_source) {

                if($attr_source != '00') {
                    if("" != trim($import_data[$model][$attr_source])) {
                        $product_attributes[] = [
                            'attribute_id' => $attr_id,
                            'product_attribute_description' => [
                                $this->defaultLang => [
                                    'text' => $import_data[$model][$attr_source]
                                ]
                            ]
                        ];
                    }
                }
            }
            /** product_attributes */

            /** product discounts */
            $discount_config = $config['eimport_activeDiscountGroup'];
            $product_discount= [];
            forEach($discount_config as $customer_group => $discount) {
                if($discount != "00") {
                    $product_discount[] = [
                        'customer_group_id' => $customer_group,
                        'priority'          => '0',
                        'price'             => $import_data[$model][$discount],
                        'date_start'        => '',
                        'date_end'          => ''
                    ];
                }
                
            }
            /** product discounts */




            /** product_option */
            $product_options = [];
            $option_values   = [];
            if(!empty($import_data[$model]['options'])) {
                
                forEach($import_data[$model]['options'] as $option) {
                    $options = explode('|', trim($option[$name_row]));
                    
                    $option_name       = $options[0];
                    $option_value_name = isset($options[1]) ? $options[1] : '';

                    $option_desc = $this->db->query("SELECT * FROM " . DB_PREFIX . "option_description WHERE language_id = '$this->defaultLang' and name = '".$this->db->escape($option_name)."'")->row;
                    
                    if(empty($option_desc)) {
                        $this->db->query("INSERT into ".DB_PREFIX."option (type, sort_order) VALUES('select', '0') ");
                        $option_id = $this->db->getLastId();

                        $this->db->query("INSERT into ".DB_PREFIX."option_description (option_id, language_id, name) VALUES('$option_id', '$this->defaultLang', '".$this->db->escape($option_name)."')");

                    }else{
                        $option_id = $option_desc['option_id'];
                    }
                    
                    $option_value_desc = $this->db->query("SELECT * FROM " . DB_PREFIX . "option_value_description WHERE option_id = '$option_id' and language_id = '$this->defaultLang' and name = '".$this->db->escape($option_value_name)."'")->row;
                    
                    if(empty($option_value_desc)) {

                        $this->db->query("INSERT into ".DB_PREFIX."option_value (option_id, image, sort_order) VALUES('$option_id', '', '0') ");
                        $option_value_id = $this->db->getLastId();

                        $this->db->query("INSERT into ".DB_PREFIX."option_value_description (option_value_id, language_id, option_id, name) VALUES('$option_value_id', '$this->defaultLang', '$option_id', '".$this->db->escape($option_value_name)."') ");


                    }else{
                        $option_value_id = $option_value_desc['option_value_id'];
                    }

                    
                    $product_options[$option_id]['type'] = 'select';
                    $product_options[$option_id]['option_id'] = $option_id;
                    $product_options[$option_id]['required'] = '1';

                    $product_options[$option_id]['product_option_value'][] = [
                        'option_value_id' => $option_value_id,
                        'quantity'        => (int) $option[$config['eimport_qty']],
                        'subtract'        => '1',
                        'price'           => $option[$config['eimport_price']] > 0 ? $option[$config['eimport_price']] : $option[$config['eimport_price']] * -1,
                        'price_prefix'    => $option[$config['eimport_price']] > 0 ? '+' : '-',
                        'points'          => '0',
                        'points_prefix'   => '+',
                        'weight'          => '0',
                        'weight_prefix'   => '+'
                    ];

                }
            }

            /** product_option */

            /** product categories */

            $row_main_category = $config['eimport_category'];
            $row_sub_category  = $config['eimport_category_sub'];
            $row_sub2_category = $config['eimport_category_sub2'];

            $product_categories = [];
            if($row_main_category != "00") {

                $category_names = $import_data[$model][$row_main_category];

                if("" != trim($category_names)) {
                    $main_categories = explode('|', trim($category_names));
                    
                    forEach($main_categories as $mk => $mc) {
                        $category_id = 0;
                        $check = $this->db->query("SELECT * FROM ".DB_PREFIX."category_description cd LEFT JOIN ".DB_PREFIX."category c ON(c.category_id = cd.category_id) WHERE c.parent_id = '0' and cd.language_id='$this->defaultLang' and cd.name='".$this->db->escape($mc)."' ")->row;

                        if(!empty($check)) {
                            $category_id = $check['category_id'];
                        }else{
                            $category_id = $this->addCategory($mc);
                        }

                        if($category_id > 0) {
                            $p_main_cats[$mk] = $category_id;
                            $product_categories[] = $category_id;
                        } 
                    }
                } /** added main category */

                if($row_sub_category != "00") {
                    $category_names = $import_data[$model][$row_sub_category];

                    if("" != trim($category_names)) {
                        $sub_categories = explode('|', trim($category_names));


                        forEach($sub_categories as $mk => $mc) {
                            $sub_category_id = 0;
                            $check = $this->db->query("SELECT * FROM ".DB_PREFIX."category_description cd  LEFT JOIN ".DB_PREFIX."category c ON(c.category_id = cd.category_id) WHERE c.parent_id = '".$p_main_cats[$mk]."' and cd.language_id='$this->defaultLang' and cd.name='".$this->db->escape($mc)."' ")->row;

                            if(!empty($check)) {
                                $sub_category_id = $check['category_id'];
                            }else{
                                $sub_category_id = $this->addCategory($mc, $p_main_cats[$mk]);
                            }

                            if($sub_category_id > 0) {
                                $p_sub_cats[$mk] = $sub_category_id;
                                $product_categories[] = $sub_category_id;
                            } 
                        }

                    }


                    if($row_sub2_category != "00") {
                        $category_names = $import_data[$model][$row_sub2_category];
    
                        if("" != trim($category_names)) {
                            $sub_categories = explode('|', trim($category_names));
    
    
                            forEach($sub_categories as $mk => $mc) {
                                $sub2_category_id = 0;
                                $check = $this->db->query("SELECT * FROM ".DB_PREFIX."category_description cd  LEFT JOIN ".DB_PREFIX."category c ON(c.category_id = cd.category_id) WHERE c.parent_id = '".$p_sub_cats[$mk]."' and cd.language_id='$this->defaultLang' and cd.name='".$this->db->escape($mc)."' ")->row;
    
                                if(!empty($check)) {
                                    $sub2_category_id = $check['category_id'];
                                }else{
                                    $sub2_category_id = $this->addCategory($mc, $p_sub_cats[$mk]);
                                }
    
                                if($sub2_category_id > 0) $product_categories[] = $sub2_category_id;
                            }
    
                        }
    
                    }

                }
            }

            $seo_url = [];

            $seo_metod = $config['eimport_seoMetod'];
            $seo_row   = $config['eimport_seoRow'];

            if($seo_metod == '1') {
                $seo_url = [
                    $this->stores[0] => [
                        $this->defaultLang => $this->createSeo($model.' '.$product_description[$this->defaultLang]['name'])
                    ]
                ];
            }elseif($seo_metod == '2' && "" != trim($import_data[$model][$seo_row]) ) {
                $seo_url = [
                    $this->stores[0] => [
                        $this->defaultLang => $this->createSeo($import_data[$model][$seo_row])
                    ]
                ];
            }elseif( $seo_metod == '3' ) {
                if($product_image != "") {
                    $parseImageName = explode('/', $product_image);
                    $parseImageName = end($parseImageName);
                    $parseImageName = explode('.', $parseImageName);

                    $seo_url = [
                        $this->stores[0] => [
                            $this->defaultLang => $this->createSeo($parseImageName[0])
                        ]
                    ];
                }else{
                    $seo_url = [
                        $this->stores[0] => [
                            $this->defaultLang => $this->createSeo($model.' '.$product_description[$this->defaultLang]['name'])
                        ]
                    ];
                }
                
            }

            $weight =$config['eimport_weight'] != "00" ? $import_data[$model][$config['eimport_weight']] : '';
            $desi = isset($config['eimport_desi']) && "00" != $config['eimport_desi'] ? trim($import_data[$model][$config['eimport_desi']]) : '';
            if( isset($config['eimport_desi_compare']) && 1 == $config['eimport_desi_compare'] && $desi != '' && $weight != '' && $desi > $weight ) $weight = $desi;
            $relateds = [];

            $relatedsRow = isset($config['eimport_related']) ? $config['eimport_related'] : "00";

            if($relatedsRow != "00") {
                $relatedSource = $import_data[$model][$relatedsRow];

                if( "" != trim($relatedSource) ) {

                    $relatedSource = explode(',', $relatedSource);

                    forEach($relatedSource as $rsId) {
                        if( 0 < (int) $rsId ) $relateds[] = (int) $rsId;
                    }

                }

            }

            /** product categories */
            $product = [
                'model' => $model,
                'sku'   => '',
                'upc'   => '',
                'ean'   => '',
                'jan'   => '',
                'isbn'  => '',
                'mpn'   => '',
                'location' => '',
                'quantity' => $config['eimport_qty'] != "00" ? $import_data[$model][$config['eimport_qty']] : '0',
                'minimum'  => '1',
                'subtract' => '1',
                'stock_status_id' => $config['eimport_stockStatus'],
                'date_available'  => date("Y-m-d"),
                'manufacturer_id' => 0,
                'shipping' => $config['eimport_reqShipping'],
                'price'    => $config['eimport_price'] != "00" ? $import_data[$model][$config['eimport_price']] : '0',
                'points'   => 0,
                'weight'   => $weight,
                'weight_class_id' => $config['eimport_weight_class'],
                'length'  => $config['eimport_dimentionL'] != "00" ? $import_data[$model][$config['eimport_dimentionL']] : '',
                'width'   => $config['eimport_dimentionW'] != "00" ? $import_data[$model][$config['eimport_dimentionW']] : '',
                'height'  => $config['eimport_dimentionH'] != "00" ? $import_data[$model][$config['eimport_dimentionH']] : '',
                'length_class_id'  => $config['eimport_lenght_class'],
                'status'  => '1',
                'tax_class_id' => $config['eimport_taxClass'],
                'sort_order'   => $i,
                'image' => $product_image,
                'product_description' => $product_description,
                'product_store' => $this->stores,
                'product_attribute' => $product_attributes,
                'product_special'  => $product_discount,
                'product_option'    => $product_options,
                'product_image'     => $product_images,
                'product_category'  => $product_categories,
                'product_seo_url'   => $seo_url,
                'product_related'   => $relateds
            ];


            /** product external fields */
            $external_names = isset($config['eimport_external_names']) ? $config['eimport_external_names'] : [];
            $external_rows = isset($config['eimport_external_rows']) ? $config['eimport_external_rows'] : [];

            if(!empty($external_names)) {

                forEach($external_names as $ek => $en) {
                    if($external_rows[$ek] != "00") {
                        $row_value = $import_data[$model][$external_rows[$ek]];
                        $product[$en] = $row_value;
                    }
                }
            }

            /** product external fields */



            $product_id = $this->isProductSaved($model);

            if($product_id == 0) $product_id = $this->saveProduct($product);
            elseif(isset($config['eimport_update_exist']) && '1' == $config['eimport_update_exist']) $product_id = $this->updateProduct($product, $product_id);
            else {
                $this->session->data['error'][] = $model.' yükle / güncelle seçeneği yok';
            }

            if($product_id == 0) {
                $this->session->data['error'][] = $model.' '.$this->language->get('err_product');
            }

            $current++;
            if($current == $limit) break;
            
        }

        $status['page'] = $i;
        
        return $status;
    }

    private function addCategory($mc, $parent_id = 0) {
        $this->load->model('catalog/category');
        $category_data = [
            'parent_id' => $parent_id,
            'top'       => 0,
            'column'    => 1,
            'sort_order'=> 0,
            'status'    => 1,
            'image'     => '',
            'category_description' => [
                $this->defaultLang => [
                    'name' => $mc,
                    'description' => '',
                    'meta_title' => $mc,
                    'meta_description' => '',
                    'meta_keyword' => ''
                ]
            ],
            'category_seo_url' => [
                $this->stores[0] => [
                    $this->defaultLang => $this->createSeo($mc)
                ]
            ],
            'category_store' => $this->stores
        ];

        $category_id = $this->model_catalog_category->addCategory($category_data);

        return $category_id;
    }

    /* takes the input, scrubs bad characters */
    private function createSeo($input, $replace = '-', $remove_words = false, $words_array = array()) {
        //make it lowercase, remove punctuation
        $utf_find = ['ü', 'ğ', 'ı', 'ş', 'ö', 'ç'];
        $utf_replace = ['u', 'g', 'i', 's', 'o', 'c'];



        $return = trim(preg_replace('/-{1,}/', '-', preg_replace('/[^a-zA-Z0-9\-\s]/', '', str_replace($utf_find, $utf_replace, mb_strtolower($input) ) ) ) );

        //remove words, if not helpful to seo
        if($remove_words) { $return = remove_words($return, $replace, $words_array); }

        //convert the spaces to whatever the user wants
        //usually a dash or underscore..
        //...then return the value.
        return str_replace(' ', $replace, $return);
    }

/* takes an input, scrubs unnecessary words */
    private function remove_words($input,$replace,$words_array = array(),$unique_words = true)
    {
        //separate all words based on spaces
        $input_array = explode(' ',$input);

        //create the return array
        $return = array();

        //loops through words, remove bad words, keep good ones
        foreach($input_array as $word)
        {
            //if it's a word we should add...
            if(!in_array($word,$words_array) && ($unique_words ? !in_array($word,$return) : true))
            {
                $return[] = $word;
            }
        }

        //return good words separated by dashes
        return implode($replace,$return);
    }

    private function getProductImages($model) {
        $images = [];

        $find   = glob("../image/catalog/{,*/,*/*/,*/*/*/}$model*.jpg", GLOB_BRACE);

        foreach($find as $sort_order => $image) {
            if( !is_dir($image) ) {
                $images[] = [
                    'image'      => str_replace('../image/catalog/', '', $image),
                    'sort_order' => $sort_order
                ];
            }
        }

        return $images;
    }

    private function isProductSaved($model) {
        $this->load->model('catalog/product');

        $search = $this->model_catalog_product->getProducts(['filter_model' => $model]);

        return empty($search) ? 0 : $search[0]['product_id'];
    }

    private function saveProduct($product) {
        $this->load->model('catalog/product');

        $product_id = $this->model_catalog_product->addProduct($product);

        return $product_id;
    }

    private function updateProduct($product, $product_id) {
        $this->load->model('catalog/product');

        forEach($product['product_option'] as $k => $v) {
            $product['product_option'][$k]['product_option_id'] = '';

            forEach($v['product_option_value'] as $kk => $vv) {
                $product['product_option'][$k]['product_option_value'][$kk]['product_option_value_id'] = '';
            }
        }

        $this->model_catalog_product->editProduct($product_id, $product);

        return $product_id;
    }
}