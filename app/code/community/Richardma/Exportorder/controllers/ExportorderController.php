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

    public function exportAction() 
    {
        $orderIds = $this->getRequest()->getPost('orderIds');
        //var_dump($this->_genOrdersIdList($orderIds));

        //$orders = Mage::getModel('sales/order')->getCollection();
        //$orders->addAttributeToFilter('id', 1);
        $orders = Mage::getModel('sales/order')->load('100000001')->getCustomerFirstname();
        var_dump($orders);
        //$orders->load();

        //foreach ($orders as $_order) {
            //var_dump($_order->getData());
        //}
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
