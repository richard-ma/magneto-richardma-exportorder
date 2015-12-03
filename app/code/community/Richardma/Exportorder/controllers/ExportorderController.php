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


    		$line = str_replace(PHP_EOL, '', $line);
    
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
        return 0;
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
