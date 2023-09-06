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
        $this->loadHelper('Excel', [
            'className' => \Fr3nch13\Excel\View\Helper\ExcelHelper::class,
        ]);
        $this->setLayout('Fr3nch13/Excel.default');
    }

    /**
     * Used to test actual layout paths being used.
     * This shouldn't be used in production.
     *
     * @return array<string, string> The list of file names to be used.
     */
    public function getFileNames(): array
    {
        return [
            'layout' => $this->_getLayoutFileName(),
            'template' => $this->_getTemplateFileName(),
        ];
    }
}
