<?php
class Richardma_Exportorder_ExportorderController extends Mage_Adminhtml_Controller_Action
{
    public function indexAction()
    {
        $this->loadLayout();

        $block = $this->getLayout()->createBlock('richardma_exportorder/template')
            ->setTemplate('form.phtml');
        $this->_addContent($block);

        $this->renderLayout();
    }

    public function exportAddressListAction() 
    {
        $orderIds = $this->getRequest()->getPost('orderIds');
        $this->_genOrdersIdList($orderIds);

        $orders = Mage::getModel('sales/order')
            ->getCollection()
            ->addAttributeToFilter('increment_id', array('in' => $orderIds))
            ->getItems();

        // export csv
	    $data = "OrderNo,Name,Address,City,Province,Post,Country,Tel\n";
        foreach ($orders as $_order) {
            $address = $_order->getShippingAddress();

		    $line = '';
            // OrderNo
            $line .= $_order->getIncrementId();
            $line .= ',';
            // Name
            $line .= $address->getData('firstname'). ' ' .$address->getData('lastname');
            $line .= ',';
            // Address
            $line .= $address->getData('street');
            $line .= ',';
            // City
            $line .= $address->getData('city');
            $line .= ',';
            // Province
            $line .= $address->getData('region');
            $line .= ',';
            // Post
            $line .= $address->getData('postcode');
            $line .= ',';
            // Country
            $line .= Mage::app()->getLocale()->getCountryTranslation($address->getData('country_id'));
            $line .= ',';
            // Tel
            $line .= $address->getData('telephone');
            $line .= ',';

            //$line = str_replace(PHP_EOL, '', $line);
    
	    	$data .= $line;
        }
        //echo $data;
        
        // Redirect output to a client’s web browser (csv)
        header('Content-Type: text/csv');
        header('Content-Disposition: attachment;filename="address-list-'.date('Y_m_d_H_i_s').'.csv"');
        header('Cache-Control: max-age=0');
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Pragma: public'); // HTTP/1.0
        
    	echo $data;
    }

    public function exportOrderAction()
    {
        $orderIds = $this->getRequest()->getPost('orderIds');
        $orderIds = $this->_genOrdersIdList($orderIds);

        $orders = Mage::getModel('sales/order')
            ->getCollection()
            ->addAttributeToFilter('increment_id', array('in' => $orderIds))
            ->getItems();

        // export Excel xml
        // Include PHPExcel
        require_once Mage::getBaseDir('lib') . "/PHPExcel/Classes/PHPExcel.php";

        //Create new PHPExcel object
        $objPHPExcel = new PHPExcel();
        
        // Set document properties
        $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
        							 ->setLastModifiedBy("Maarten Balliauw")
        							 ->setTitle("Office 2007 XLSX Test Document")
        							 ->setSubject("Office 2007 XLSX Test Document")
        							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
        							 ->setKeywords("office 2007 openxml php")
        							 ->setCategory("Test result file");
        
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(40);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);

        // Add some data
        $delta = 6;
        $start = 1;
        foreach ($orders as $_order) {
          $address = $_order->getShippingAddress();
          $items = $_order->getAllItems();
          $end = $start + $delta - 1;
          $objPHPExcel->getActiveSheet()
                      ->mergeCells('A'.$start.':A'.$end.'')
                      ->setCellValue('A'.$start.'', $_order->getIncrementId())

                      ->setCellValue('B'.$start.'', 'size: ' . $items[0]->getProductOptions()['options'][0]['value'])
                      ->mergeCells('B'.(string)($start+1).':B'.$end.'')

                      ->setCellValue('C'.$start.'', $address->getData('firstname'). ' ' .$address->getData('lastname'))
                      ->setCellValue('C'.(string)($start + 1).'', $address->getData('street'))
                      ->setCellValue('C'.(string)($start + 2).'', $address->getData('region'))
                      ->setCellValue('C'.(string)($start + 3).'', $address->getData('city'). ',' .$address->getData('region'). ' ' .$address->getData('postcode'))
                      ->setCellValue('C'.(string)($start + 4).'', Mage::app()->getLocale()->getCountryTranslation($address->getData('country_id')))
                      ->setCellValue('C'.(string)($start + 5).'', $address->getData('telephone'))
                      ->getStyle('C'.(string)($start + 5))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
          // add picture
          $objDrawing = new PHPExcel_Worksheet_Drawing();
          $imageUrl = $items[0]->getProduct()->getImageUrl();
          $imageUrl = str_replace(Mage::getBaseUrl('media'), Mage::getBaseDir('media').'/', $imageUrl);
          $objDrawing->setPath($imageUrl);
          $objDrawing->setCoordinates('B'.(string)($start+1));
          $objDrawing->setHeight(80);
          $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());

          $start = $start + $delta;
        }
        
        // Rename worksheet
        $objPHPExcel->getActiveSheet()->setTitle('Orders');
        
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="orders-'.date('Y_m_d_H_i_s').'.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
    }

    private function _genOrdersIdList($post_string) 
    {
        $temp_list = split(',', $post_string);
        $ret_list = array();
        foreach($temp_list as $id) {
            $id = trim($id);
            if (strpos($id, '-') != false) {
                $range = split('-', $id);
                $start = $range[0];
                $end = $range[1];
                for ($i = $start; $i <= $end; $i++) {
                    array_push($ret_list, $i);
                }
            } else {
                array_push($ret_list, $id);
            }
        }
        $ret_list = array_unique($ret_list);
        sort($ret_list);

        return $ret_list;
    }
}
