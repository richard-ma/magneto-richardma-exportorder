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
        $this->loadLayout();

        $block = $this->getLayout()->createBlock('richardma_exportorder/template')
            ->setTemplate('export.phtml');
        $this->_addContent($block);

        $this->renderLayout();
    }
}
