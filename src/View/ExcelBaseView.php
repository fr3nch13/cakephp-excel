<?php
declare(strict_types=1);

/**
 * ExcelView
 */

namespace Fr3nch13\Excel\View;

/**
 * Excel View
 *
 * @property \Fr3nch13\Excel\View\Helper\ExcelHelper $Excel
 */
class ExcelBaseView extends \App\View\AppView
{
    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();
        $this->loadHelper('Excel', ['className' => 'Fr3nch13/Excel.Excel']);
        $this->setLayout('Fr3nch13/Excel.default');
    }
}
