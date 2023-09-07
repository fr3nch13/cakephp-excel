<?php
declare(strict_types=1);

/**
 * ExcelBaseView
 */

namespace Fr3nch13\Excel\View;

/**
 * Excel Base View
 *
 * @property \Fr3nch13\Excel\View\Helper\ExcelHelper $Excel
 */
class ExcelBaseView extends \App\View\AppView
{
    /**
     * @var string The path to look for the layout.
     */
    protected $layoutPath = '';

    /**
     * @var string The sub directory to look for the template.
     */
    protected $subDir = '';

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
