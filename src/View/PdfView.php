<?php
declare(strict_types=1);

/**
 * PdfView
 */

namespace Fr3nch13\Excel\View;

/**
 * PDF View
 */
class PdfView extends ExcelBaseView
{
    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->setLayout('Fr3nch13/Excel.pdf/default');

        $this->getResponse()->setTypeMap('pdf', ['application/pdf']);
        $this->setResponse($this->getResponse()->withType('pdf'));
    }
}
