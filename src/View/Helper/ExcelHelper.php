<?php
declare(strict_types=1);

/**
 * ExcelHelper
 */

namespace Fr3nch13\Excel\View\Helper;

use Fr3nch13\Excel\Excel;

/**
 * Excel Helper
 *
 * Helper write Excel files to the brouswer output.
 * Essentially just a wrapper to \Fr3nch13\Excel\Excel
 */
class ExcelHelper extends \Cake\View\Helper
{
    /**
     * @var array<string, mixed> $config The configuration options.
     */
    public $config = [];

    /**
     * @var \Fr3nch13\Excel\Excel The excel object.
     */
    public $Excel;

    /**
     * Constructor hook method.
     *
     * @param array<string, mixed> $config The configuration options.
     * @return void
     */
    public function initialize(array $config = []): void
    {
        $this->config = $config;
        $this->Excel = new Excel();
    }

    /**
     * Wrapper for the \Fr3nch13\Excel\Excel::create() method.
     *
     * @param array<string, mixed> $properties The properties for the excel object/file.
     * @param array<string, string> $headers The first line of the excel file, also defines the column mapping.
     * @param array<int, mixed> $rows The rest of the rows.
     * @return \Fr3nch13\Excel\Excel The created excel object With the supplied data.
     */
    public function create(array $properties = [], array $headers = [], array $rows = []): Excel
    {
        return $this->Excel->create($properties, $headers, $rows);
    }
}
