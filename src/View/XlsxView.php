<?php
declare(strict_types=1);

/**
 * XlsxView
 */

namespace Fr3nch13\Excel\View;

/**
 * Xlsx View
 */
class XlsxView extends ExcelBaseView
{
    /**
     * @var string The path to look for the layout.
     */
    protected $layoutPath = 'xlsx';

    /**
     * @var string The sub directory to look for the template.
     */
    protected $subDir = 'xlsx';

    /**
     * Mime-type this view class renders as.
     *
     * @return string The JSON content type.
     */
    public static function contentType(): string
    {
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->getResponse()->setTypeMap('xlsx', ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);
        $this->setResponse($this->getResponse()->withType('xlsx'));
    }
}
