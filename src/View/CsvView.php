<?php
declare(strict_types=1);

/**
 * CsvView
 */

namespace Fr3nch13\Excel\View;

/**
 * CSV View
 */
class CsvView extends ExcelBaseView
{
    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->setLayout('Fr3nch13/Excel.csv/default');

        $this->getResponse()->setTypeMap('csv', ['text/csv; charset=UTF-8']);
        $this->setResponse($this->getResponse()->withType('csv'));
    }
}
