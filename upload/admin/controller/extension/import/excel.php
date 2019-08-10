<?php
class ControllerExtensionImportExcel extends Controller {
    public function index() {
        $this->load->language('extension/excel/excel');

        $this->document->setTitle($this->language->get('heading_title'));
        
        if(isset($this->request->post['eimport_model'])) {
			$this->load->model('setting/setting');
			$this->model_setting_setting->editSetting('eimport', $this->request->post);
        }

        $this->getForm();
	}
	
	public function upload() {
		$this->load->language('extension/extension/excel');
		if (!empty($this->request->files['excel_file']['name']) && is_file($this->request->files['excel_file']['tmp_name'])) {
			$import_file = html_entity_decode($this->request->files['excel_file']['name'], ENT_QUOTES, 'UTF-8');
			if( !file_exists(DIR_UPLOAD.'/excel_import') ) mkdir(DIR_UPLOAD.'/excel_import');
			$import_file = 'excel_import/'.$import_file;
			
			$link = $this->url->link('extension/import/excel', '', TRUE).'&user_token=' . $this->session->data['user_token'];
			if( move_uploaded_file($this->request->files['excel_file']['tmp_name'], DIR_UPLOAD . $import_file) ) {
				$link = $this->url->link('extension/import/excel/import', '', TRUE).'&filename='. $import_file .'&page=0&user_token='.$this->session->data['user_token'];
			}else{
				$this->session->data['error'] = $this->language->get('error_upload');
			}
		}else{
			$this->session->data['error'] = $this->language->get('error_no_file');
		}

		echo 'Please wait until <a href="'.$link.'">redirect</a>';
		echo "<script>window.location='{$link}'</script>";
	}

	public function import() {
		$this->load->language('extension/excel/excel');
		$this->load->model('extension/excel');
		
		$page = isset($this->request->get['page']) ? (int) $this->request->get['page'] : 0;
		$import_file= $this->request->get['filename'];

		$import_status = $this->model_extension_excel->import($import_file, $page);
		

		$link = $this->url->link('extension/import/excel', '', TRUE).'&user_token=' . $this->session->data['user_token'];
		if($import_status['total'] == $import_status['page']) {
			$this->session->data['success'] = $this->language->get('import_success');
		}else{
			$link = $this->url->link('extension/import/excel/import', '', TRUE).'&filename='.$import_file.'&page='.$import_status['page'].'&user_token='.$this->session->data['user_token'];
		}


		echo 'Please Wait until redirect to <a href="'.$link.'">next step</a>';
		//echo '<script>window.location="'.$link.'";</script>';
	}

    protected function getForm() {
        $data['breadcrumbs'] = array();

		$data['breadcrumbs'][] = array(
			'text' => $this->language->get('text_home'),
			'href' => $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'], true)
		);

		$data['breadcrumbs'][] = array(
			'text' => $this->language->get('heading_title'),
			'href' => $this->url->link('extension/import/excel', 'user_token=' . $this->session->data['user_token'], true)
		);

		$data['cancel'] = $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'], true);
		$data['action'] = $this->url->link('extension/import/excel', 'user_token=' . $this->session->data['user_token'], true);
		$data['upload_action'] = $this->url->link('extension/import/excel/upload', 'user_token=' . $this->session->data['user_token'], true);
        
        $temp_rows = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
        $rows = $temp_rows;

        foreach($temp_rows as $row) {$rows[] = 'A'.$row;}
        foreach($temp_rows as $row) {$rows[] = 'B'.$row;}

        unset($temp_rows);
        $data['rows'] = $rows;

		$this->load->model('setting/setting');
		
        $data['excel_data'] = $this->model_setting_setting->getSetting('eimport');
        $this->load->model('localisation/tax_class');
		$data['tax_classes'] = $this->model_localisation_tax_class->getTaxClasses();

        $this->load->model('localisation/stock_status');
		$data['stock_statuses'] = $this->model_localisation_stock_status->getStockStatuses();

        $this->load->model('localisation/weight_class');
        $data['weight_classes'] = $this->model_localisation_weight_class->getWeightClasses();
        
        $this->load->model('localisation/length_class');
		$data['length_classes'] = $this->model_localisation_length_class->getLengthClasses();


        // Attributes
		$this->load->model('catalog/attribute');

		$data['attributes'] = array();

		$filter_data = array(
			'start' => 0,
			'limit' => 500
		);

		$attributes = $this->model_catalog_attribute->getAttributes($filter_data);

		foreach ($attributes as $result) {
			$data['attributes'][] = array(
				'attribute_id'    => $result['attribute_id'],
				'name'            => $result['name'],
				'attribute_group' => $result['attribute_group'],
				'sort_order'      => $result['sort_order']
			);
        }
        
        $this->load->model('customer/customer_group');

		$data['customer_groups'] = $this->model_customer_customer_group->getCustomerGroups();


		if(isset($this->session->data['error'])) {
			$data['error_warning'] = is_array($this->session->data['error']) ? implode('<br />', $this->session->data['error']) : $this->session->data['error'];
			unset($this->session->data['error']);
		}

		if(isset($this->session->data['success'])) {
			$data['success'] = $this->session->data['success'];
			unset($this->session->data['success']);
		}

        $data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');


        $this->response->setOutput($this->load->view('extension/extension/excel', $data));
    }
}