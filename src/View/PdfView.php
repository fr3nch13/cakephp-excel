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
     * @var string The path to look for the layout.
     */
    protected $layoutPath = 'pdf';

    /**
     * @var string The sub directory to look for the template.
     */
    protected $subDir = 'pdf';

    /**
     * Mime-type this view class renders as.
     *
     * @return string The JSON content type.
     */
    public static function contentType(): string
    {
        return 'application/pdf';
    }

    /**
     * Initialize method
     *
     * @return void
     */
    public function initialize(): void
    {
        parent::initialize();

        $this->setLayoutPath('pdf');
        $this->setSubDir('pdf');

        $this->getResponse()->setTypeMap('pdf', [$this->contentType()]);
        $this->setResponse($this->getResponse()->withType('pdf'));
    }
}
