<?php
declare(strict_types=1);

/**
 * ExcelView
 */

namespace Fr3nch13\Excel\View;

/**
 * Excel View
 */
class ExcelView extends ExcelBaseView
{
    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->setLayout('Fr3nch13/Excel.xlsx/default');
        $this->setSubDir('xlsx');

        $this->getResponse()->setTypeMap('xlsx', ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);
        $this->setResponse($this->getResponse()->withType('xlsx'));
    }
}
