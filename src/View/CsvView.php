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
     * @var string The path to look for the layout.
     */
    protected $layoutPath = 'csv';

    /**
     * @var string The sub directory to look for the template.
     */
    protected $subDir = 'csv';

    /**
     * Mime-type this view class renders as.
     *
     * @return string The JSON content type.
     */
    public static function contentType(): string
    {
        return 'text/csv; charset=UTF-8';
    }

    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->getResponse()->setTypeMap('csv', ['text/csv; charset=UTF-8']);
        $this->setResponse($this->getResponse()->withType('csv'));
    }
}
