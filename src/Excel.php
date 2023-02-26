<?php
declare(strict_types=1);

/**
 * Excel
 */

namespace Fr3nch13\Excel;

use Cake\Core\Configure;
use Cake\Log\Log;
use Cake\Utility\Text;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Writer\Csv as CsvWriter;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf as PdfWriter;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;

/**
 * Excel
 *
 * CakePHP friendly wrapper for PHPExcel
 *
 * @link https://github.com/PHPOffice/PhpSpreadsheet
 */
class Excel
{
    /**
     * The internal data that the excel file is parsed into.
     *
     * @var array<string, mixed>
     */
    protected $data = [];

    /**
     * The list of possible errors.
     *
     * @var array<int, string>
     */
    protected $errors = [];

    /**
     * The absolute path to the excel file.
     *
     * @var string
     */
    protected $filePath;

    /**
     * The type of file we're reading, usually excel or csv.
     *
     * @var null|string
     */
    protected $inputFileType = null;

    /**
     * The max amount of lines a reader will return.
     *
     * @var null|int
     */
    protected $lineLimit = null;

    /**
     * The properties of the spreadsheet.
     *
     * @var array<string, mixed>
     */
    protected $properties = [];

    /**
     * The reader object instance created using the IO Factory.
     *
     * @var \PhpOffice\PhpSpreadsheet\Reader\IReader|null
     */
    protected $reader = null;

    /**
     * If we are reading, or writing.
     *
     * @var bool
     */
    protected $readwrite;

    /**
     * The sheet index number we're scoped to.
     *
     * @var null|int
     */
    protected $sheetIndex = null;

    /**
     * The spreadsheet object instance created using the IO Factory.
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet|null
     */
    protected $spreadsheet = null;

    /**
     * The writer object instance created using the IO Factory.
     *
     * @var \PhpOffice\PhpSpreadsheet\Writer\IWriter
     */
    protected $writer;

    /**
     * The memory tracking object.
     *
     * @var null|\Fr3nch13\Excel\Memory;
     */
    protected $Memory = null;

    /**
     * The initilization method
     *
     * @param null|string $filePath The absolute path to the excel file.
     * @param int $sheetIndex The specific sheet we're working with by it's index number.
     *      If we're writing, we'll write only to this sheet index.
     *      If we're reading, we'll only return the results from this sheet.
     * @param bool $write false if we're reading, true if we're writing
     * @param null|int $lineLimit Limit the amount of returned lines an excel file will be using.
     */
    public function __construct(
        ?string $filePath = null,
        int $sheetIndex = 0,
        bool $write = false,
        ?int $lineLimit = null
    ) {
        set_error_handler(function ($severity, $message, $filename, $lineno) {
            if (error_reporting() == 0) {
                return true;
            }
            if (error_reporting() & $severity) {
                throw new \ErrorException($message, 0, $severity, $filename, $lineno);
            }
        });

        if ($filePath) {
            $this->setFilePath($filePath);
        }
        $this->setSheetIndex($sheetIndex);
        $this->readwrite = $write;

        if ($lineLimit) {
            $this->setLineLimit($lineLimit);
        }

        $this->errors = [];
    }

    /**
     * Mainly used to restor the error handler.
     *
     * @return void
     */
    public function __destruct()
    {
        restore_error_handler();
    }

    /**
     * Reports the memory usage at the time it is called.
     *
     * @param bool $nice If we should return the bytes (false), of the calculated amount in a nice format (true).
     * @param float|null $mem_usage The memory number to be made nice.
     * @return string the memory usage stat.
     */
    public function memoryUsage(bool $nice = true, ?float $mem_usage = null): string
    {
        if (!$this->Memory) {
            $this->Memory = new Memory();
        }

        return $this->Memory->memoryUsage($nice, $mem_usage);
    }

    /**
     * Sets $this->lineLimit
     *
     * @param null|int $lineLimit The limit of lines that we'll return.
     * @return \Fr3nch13\Excel\Excel The instance that we're creating to allow chaining.
     */
    public function setLineLimit(?int $lineLimit = null): \Fr3nch13\Excel\Excel
    {
        $this->lineLimit = $lineLimit;

        return $this;
    }

    /**
     * Gets the line limit
     *
     * @return null|int the value of $this->filePath
     */
    public function getLineLimit(): ?int
    {
        return $this->lineLimit;
    }

    /**
     * Sets $this->filePath
     *
     * @param string $filePath The absolute path to the excel file.
     * @return \Fr3nch13\Excel\Excel The instance that we're creating to allow chaining.
     */
    public function setFilePath(string $filePath): \Fr3nch13\Excel\Excel
    {
        $this->filePath = $filePath;

        return $this;
    }

    /**
     * Gets the file path
     *
     * @return string the value of $this->filePath
     */
    public function getFilePath(): string
    {
        return $this->filePath;
    }

    /**
     * Sets $this->sheetIndex
     *
     * @param int $sheetIndex The index number to the sheet we want to work with
     * @return \Fr3nch13\Excel\Excel The instance that we're creating to allow chaining.
     */
    public function setSheetIndex(int $sheetIndex): \Fr3nch13\Excel\Excel
    {
        $this->sheetIndex = $sheetIndex;

        return $this;
    }

    /**
     * Gets the sheet index
     *
     * @return null|int the value of $this->sheetIndex
     */
    public function getSheetIndex(): ?int
    {
        return $this->sheetIndex;
    }

    /**
     * Reads excel file, and parses it into an array for internal uses.
     *
     * @param null|\Cake\Console\ConsoleIo $io The ConsoleIo for writing out info.
     * @return \Fr3nch13\Excel\Excel|null The instance that we're creating to allow chaining.
     * @TODO instead of returning null, allow the exceptions to go through, or create aspecific one, and throw it.
     */
    public function load(?\Cake\Console\ConsoleIo $io = null): ?\Fr3nch13\Excel\Excel
    {
        $start = time();
        Log::write('debug', __('Loading spreadsheet (This may take awhile) ...'));

        try {
            \PhpOffice\PhpSpreadsheet\Settings::setLocale('en_us');
            $this->inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($this->getFilePath());
            $this->reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($this->inputFileType);
            $sheetIndex = $this->getSheetIndex();
            $this->spreadsheet = $this->getReader()->load($this->getFilePath());
            if ($sheetIndex === null) {
                $this->getSpreadsheet()->setActiveSheetIndex(0);
            } else {
                $this->getSpreadsheet()->setActiveSheetIndex($sheetIndex);
            }
        } catch (\ErrorException $e) {
            $this->setError(__('Error when loading excel file: {0}', [$e->getMessage()]));

            return null;
        }
        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }

        Log::write('debug', __('Loaded and created spreadsheet read object in `{0}` seconds. Memory:`{1}`', [
            time() - $start,
            $mem_usage,
        ]));

        try {
            //get the properties
            $this->data['properties'] = $this->getProperties();

            // get the info
            /** @var \PhpOffice\PhpSpreadsheet\Reader\Xlsx $reader */
            $reader = $this->getReader();
            $this->data['info'] = $reader->listWorksheetInfo($this->getFilePath());

            // get the sheet names
            $this->data['sheets'] = $this->getSheetNames();
        } catch (\ErrorException $e) {
            $this->setError(__('Error when loading excel file: {0}', [$e->getMessage()]));

            return null;
        }

        return $this;
    }

    /**
     * This creates the write excel write object, and sets initial properties.
     *
     * @param array<string, string> $properties The properties for the excel object/file.
     * @param array<string, string> $headers The first line of the excel file, also defines the column mapping.
     * @param array<int, mixed> $rows The rest of the rows.
     * @return \Fr3nch13\Excel\Excel The created excel object With the supplied data.
     */
    public function create(array $properties = [], array $headers = [], array $rows = []): \Fr3nch13\Excel\Excel
    {
        \PhpOffice\PhpSpreadsheet\Settings::setLocale('en_us');
        $this->spreadsheet = new Spreadsheet();
        $this->setSheetIndex(0);
        $this->getSpreadsheet()->setActiveSheetIndex($this->getSheetIndex());
        if (isset($properties['creator'])) {
            $this->getSpreadsheet()->getProperties()->setCreator($properties['creator']);
        }
        if (isset($properties['modifier'])) {
            $this->getSpreadsheet()->getProperties()->setLastModifiedBy($properties['modifier']);
        }
        if (isset($properties['title'])) {
            $this->getSpreadsheet()->getProperties()->setTitle($properties['title']);
        }
        if (isset($properties['subject'])) {
            $this->getSpreadsheet()->getProperties()->setSubject($properties['subject']);
        }
        if (isset($properties['description'])) {
            $this->getSpreadsheet()->getProperties()->setSubject($properties['description']);
        }
        if (isset($properties['keywords'])) {
            $this->getSpreadsheet()->getProperties()->setKeywords($properties['keywords']);
        }
        if (isset($properties['cetegory'])) {
            $this->getSpreadsheet()->getProperties()->setCategory($properties['cetegory']);
        }

        $sheet = $this->getSpreadsheet()->getActiveSheet();

        if (isset($properties['sheet_title'])) {
            $sheet->setTitle($properties['sheet_title']);
        }

        $rowNum = 1; // row num
        $colLetter = 'A';
        $headerMap = [];
        foreach ($headers as $key => $value) {
            $headerMap[$key] = $colLetter;
            $sheet->setCellValue($headerMap[$key] . $rowNum, $value);
            $colLetter++;
        }

        foreach ($rows as $i => $cells) {
            $rowNum++;
            foreach ($cells as $key => $value) {
                $colNum = $headerMap[$key];
                $sheet->setCellValue($colNum . $rowNum, $value);
            }
        }

        // after writing all of the rows, auto-size the column widths.
        foreach ($headerMap as $key => $colNum) {
            $this->getSpreadsheet()->getActiveSheet()->getColumnDimension($colNum)->setAutoSize(true);
        }

        return $this;
    }

    /**
     * Tell the browser to download the compiled spreadsheet.
     *
     * @param null|string $type The type of download to create. Defaults to Xlsx.
     * @return void
     */
    public function download(?string $type = 'Xlsx'): void
    {
        if ($type == 'Xlsx') {
            $this->writer = new XlsxWriter($this->getSpreadsheet());
            $this->getWriter()->setPreCalculateFormulas(false);
        } elseif ($type == 'Pdf') {
            $this->writer = new PdfWriter($this->getSpreadsheet());
            $this->getWriter()->setPreCalculateFormulas(false);
        } elseif ($type == 'Csv') {
            $this->writer = new CsvWriter($this->getSpreadsheet());
            $this->getWriter()->setPreCalculateFormulas(false);
        }

        if ($this->getWriter()) {
            $this->getWriter()->save('php://output');
        }
    }

    /**
     * Save the modified spreadsheet.
     *
     * @return bool True if saved, false if it failed for some reason. If it failed, check the errors with getErrors();
     */
    public function save(): bool
    {
        $start = time();
        $sheetIndex = $this->getSheetIndex();
        if (!$sheetIndex) {
            $sheetIndex = 0;
        }

        // set some of the properties
        $this->getSpreadsheet()->setActiveSheetIndex($sheetIndex);
        $this->getSpreadsheet()->getActiveSheet()->getProtection()->setSheet(true);
        $this->getSpreadsheet()->getProperties()
            ->setLastModifiedBy(Configure::read('Theme.title'));
        $this->writer = new XlsxWriter($this->getSpreadsheet());
        $this->getWriter()->setPreCalculateFormulas(false);
        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        Log::write('debug', __('Created spreadsheet write object. Memory:`{0}`', [$mem_usage]));

        // Temporary to check generated sheet against the original one.
        $filePath = $this->getFilePath();

        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        Log::write('debug', __('Saving updated spreadsheet to `{0}`. Memory:`{1}`', [
            $filePath, $mem_usage,
        ]));

        try {
            $this->getWriter()->save($filePath);

            $mem_usage = memory_get_usage(true);
            if ($mem_usage < 1024) {
                $mem_usage = $mem_usage . ' B';
            } elseif ($mem_usage < 1048576) {
                $mem_usage = round($mem_usage / 1024, 2) . ' KB';
            } elseif ($mem_usage < 1073741824) {
                $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
            } else {
                $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
            }
            Log::write('debug', __('Saved updated shreadsheet to `{0}` in `{1}` seconds. Memory:`{2}`', [
                $filePath, time() - $start, $mem_usage,
            ]));
        } catch (\Throwable $e) {
            $mem_usage = memory_get_usage(true);
            if ($mem_usage < 1024) {
                $mem_usage = $mem_usage . ' B';
            } elseif ($mem_usage < 1048576) {
                $mem_usage = round($mem_usage / 1024, 2) . ' KB';
            } elseif ($mem_usage < 1073741824) {
                $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
            } else {
                $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
            }
            $msg = __('Unable to save the spreadsheet to `{0}`, Error:`{1}`, Memory:`{2}`', [
                $filePath, $e->getMessage(), $mem_usage,
            ]);
            $this->setError($msg);

            return false;
        }

        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        Log::write('debug', __('Saved updated shreadsheet to `{0}` in `{1}` seconds. Memory:`{2}`', [
            $filePath, time() - $start, $mem_usage,
        ]));

        return true;
    }

    /**
     * Gets the current reader object, and returns it
     *
     * @return \PhpOffice\PhpSpreadsheet\Reader\IReader The reader object
     */
    public function getReader(): \PhpOffice\PhpSpreadsheet\Reader\IReader
    {
        if (!$this->reader) {
            $this->load();
        }

        return $this->reader;
    }

    /**
     * Gets the current writer object, and returns it
     *
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter The writer object
     */
    public function getWriter(): ?\PhpOffice\PhpSpreadsheet\Writer\IWriter
    {
        return $this->writer;
    }

    /**
     * Gets the current spreadsheet object, and returns it
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet The spreadsheet object
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if (!$this->spreadsheet) {
            $this->load();
        }

        return $this->spreadsheet;
    }

    /**
     * Get the properties if the current spreadsheet
     *
     * @param null|string $property The name of the specific property that we need.
     * @return array<string, mixed>|string|bool The value of the properties.
     */
    public function getProperties(?string $property = null)
    {
        if (!$this->properties) {
            $this->properties = (array)$this->getSpreadsheet()->getProperties();
        }

        if ($property) {
            if (isset($this->properties[$property])) {
                return $this->properties[$property];
            } else {
                return false;
            }
        }

        return $this->properties;
    }

    /**
     * Gets the list of sheets and their names
     *
     * @return array<int, string> The list of spreadsheet names, and their index
     */
    public function getSheetNames(): array
    {
        if (!$this->spreadsheet) {
            $this->load();
        }

        return $this->getSpreadsheet()->getSheetNames();
    }

    /**
     * Sets the errors in the error array
     *
     * @param string $msg The error Message.
     * @return void
     */
    public function setError(string $msg): void
    {
        $this->errors[] = $msg;
        Log::write('error', $msg);
    }

    /**
     * Gets the current error list.
     *
     * @return array<int, string> The list of errors
     */
    public function getErrors(): array
    {
        return $this->errors;
    }

    /**
     * Update the spreadsheet from an array.
     *
     * @param array<int, array<mixed>> $rows The rows that need to be updated. Format: [(row number) => [(cell location like A/B/C) => ['value' => (value), 'options' => (array of options)]]]
     * @return bool True if everything was added as expected.
     * @throws \Throwable If an error happens in the underlying library.
     */
    public function updateFromArray(array $rows = []): bool
    {
        $start = time();
        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        $highestRow = count($rows);
        $cnt = 0;
        Log::write('debug', __('Updating `{0}` rows. Memory:`{1}`', [
            $highestRow, $mem_usage,
        ]));
        foreach ($rows as $rowNum => $row) {
            $cnt++;
            if ($cnt % 20 == 0) {
                if (Configure::read('debug')) {
                    $mem_usage = memory_get_usage(true);
                    if ($mem_usage < 1024) {
                        $mem_usage = $mem_usage . ' B';
                    } elseif ($mem_usage < 1048576) {
                        $mem_usage = round($mem_usage / 1024, 2) . ' KB';
                    } elseif ($mem_usage < 1073741824) {
                        $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
                    } else {
                        $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
                    }
                    Log::write('debug', __('Updating spreadsheet at `{0}`/`{1}` in `{2}` seconds. Memory:`{3}`', [
                        $cnt, $highestRow, time() - $start, $mem_usage,
                    ]));
                }
            }
            foreach ($row as $cellNum => $cell) {
                $cellIdx = "{$cellNum}{$rowNum}";
                $cellVal = null;
                if (is_array($cell)) {
                    if (!isset($cell['value'])) {
                        continue;
                    }
                    if (isset($cell['options'])) {
                        if (isset($cell['options']['lock'])) {
                            if ($cell['options']['lock']) {
                                $this->getSpreadsheet()->getActiveSheet()
                                    ->getStyle($cellIdx)
                                    ->getProtection()->setLocked(
                                        Protection::PROTECTION_PROTECTED
                                    );
                            } else {
                                $this->getSpreadsheet()->getActiveSheet()
                                    ->getStyle($cellIdx)
                                    ->getProtection()->setLocked(
                                        Protection::PROTECTION_UNPROTECTED
                                    );
                            }
                        }
                    }
                    $cellVal = $cell['value'];
                } else {
                    $cellVal = $cell;
                }

                if (is_object($cellVal)) {
                    if ($cellVal instanceof \Cake\I18n\FrozenTime) {
                        $cellVal = $cellVal->toUnixString();
                        try {
                            $this->getSpreadsheet()
                                ->getActiveSheet()
                                ->setCellValue($cellIdx, Date::PHPToExcel($cellVal));
                        } catch (\Throwable $e) {
                            $msg = __('1 - Issue with cell value of `{0}`', [
                                $cellVal,
                            ]);
                            $this->setError($msg);
                            throw $e;
                        }
                        $this->getSpreadsheet()->getActiveSheet()->getStyle($cellIdx)
                            ->getNumberFormat()
                            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);

                        continue;
                    }
                }

                try {
                    $this->getSpreadsheet()->getActiveSheet()->setCellValue($cellIdx, (string)$cellVal);
                } catch (\Throwable $e) {
                    $this->setError(__('2 - Issue with cell value of `{0}`', [
                        $cellVal,
                    ]));
                    throw $e;
                }
            }
        }
        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        Log::write('debug', __('Finished updating `{0}` rows in `{1}` seconds. Memory:`{2}`', [
            $highestRow, time() - $start, $mem_usage,
        ]));

        return true;
    }

    /**
     * Loads the spreadsheet into the data array
     *
     * @param null|int $sheetIndex Convert only these sheets to an array
     *      If null, it'll get the sheet names, and iterate over that
     *      If $this->sheetNum is set, then that should be the only sheet listed
     * @param null|\Cake\Console\ConsoleIo $io The ConsoleIo for writing out info.
     * @return array<string, array<mixed>> The array created from the spreadsheet
     */
    public function toArray($sheetIndex = null, ?\Cake\Console\ConsoleIo $io = null): array
    {
        Log::write('debug', __('Loading spreadsheet into an array.'));
        $sheetNames = $this->getSpreadsheet()->getSheetNames();

        // this sheet doesn't exist
        if ($sheetIndex !== null) {
            if (!isset($sheetNames[$sheetIndex])) {
                return [];
            }
            //only read this sheet
            $sheetNames = [
                $sheetIndex => $sheetNames[$sheetIndex],
            ];
        }

        $this->data['sheet_data'] = [];
        $mem_usage = memory_get_usage(true);
        if ($mem_usage < 1024) {
            $mem_usage = $mem_usage . ' B';
        } elseif ($mem_usage < 1048576) {
            $mem_usage = round($mem_usage / 1024, 2) . ' KB';
        } elseif ($mem_usage < 1073741824) {
            $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
        } else {
            $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
        }
        Log::write('debug', __('Reading File:`{0}`. Memory:`{1}`', [
            $this->getFilePath(), $mem_usage,
        ]));

        $lineLimit = $this->getLineLimit();

        foreach ($sheetNames as $i => $sheetName) {
            Log::write('debug', __('Reading Sheet:`{0} ({1})`', [
                $sheetName, $i,
            ]));
            $this->getSpreadsheet()->setActiveSheetIndex($i);
            $worksheet = $this->getSpreadsheet()->getActiveSheet();
            $highestRow = $worksheet->getHighestRow();
            Log::write('debug', __('Reporting `{0}` rows for Sheet:`{1} ({2})`', [
                $highestRow, $sheetName, $i,
            ]));
            if ($lineLimit && $lineLimit < $highestRow) {
                $highestRow = $lineLimit;
                Log::write('warning', __('Linelimit is set, only processing the first `{0}` lines for:`{1} ({2})`. ' .
                    'This is most likely caused by someone placing a character very far down by accident.', [
                    $highestRow,
                    $sheetName,
                    $i,
                ]));
            }
            $cnt = 0;
            $lstart = microtime(true);

            $progress = null;
            if ($io && $io->level() == \Cake\Console\ConsoleIo::VERBOSE) {
                /** @var \Fr3nch13\Utilities\Shell\Helper\ProgressinfoHelper $progress */
                $progress = $io->helper('Fr3nch13/Utilities.Progressinfo');
                $progress->init([
                    'total' => $highestRow,
                    'showcount' => true,
                ]);
            }

            foreach ($worksheet->getRowIterator() as $j => $row) {
                $cnt++;
                // stop at the highest limit.
                if ($cnt > $highestRow) {
                    break;
                }
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
                foreach ($cellIterator as $k => $cell) {
                    $this->data['sheet_data'][$i][$j][$k] = '';
                    try {
                        $this->data['sheet_data'][$i][$j][$k] = $cell->getFormattedValue();
                    } catch (\Throwable $e) {
                        Log::write('warning', __('Cell read error 1: File:`{0}`, Sheet:`{1}`' .
                            ', Row:`{2}`, Column:`{3}`, Msg:`{4}`', [
                            $this->getFilePath(),
                            $sheetName,
                            $j, $k,
                            $e->getMessage(),
                        ]));
                        try {
                            $this->data['sheet_data'][$i][$j][$k] = $cell->getValue();
                        } catch (\Throwable $e) {
                            $this->data['sheet_data'][$i][$j][$k] = '';
                            Log::write('warning', __('Cell read error 2: File:`{0}`, Sheet:`{1}`' .
                                ', Row:`{2}`, Column:`{3}`, Msg:`{4}`', [
                                $this->getFilePath(),
                                $sheetName,
                                $j, $k,
                                $e->getMessage(),
                            ]));
                        }
                    }
                }
                // check if the row is empty, if so then remove it.
                $rowEmpty = true;
                foreach ($this->data['sheet_data'][$i][$j] as $k => $v) {
                    if (trim($v)) {
                        $rowEmpty = false;
                        break;
                    }
                }
                if ($rowEmpty) {
                    unset($this->data['sheet_data'][$i][$j]);
                }

                if ($progress) {
                    $progress->increment(1);
                    $progress->draw(__('Memory:`{0}`', [
                        $this->memoryUsage(),
                    ]));
                }
            }

            if ($io) {
                $io->verbose(' ');
            }

            Log::write('debug', __('Found `{0}` rows of `{1}` for Sheet:`{2} ({3})` in `{4}` seconds. Memory:`{5}`', [
                count($this->data['sheet_data'][$i]),
                $highestRow,
                $sheetName,
                $i,
                microtime(true) - $lstart,
                $this->memoryUsage(),
            ]));
        }

        return $this->data;
    }

    /**
     * Turns a csv string into an array
     *
     * @param string $csvString The csv formatted string to be converted
     * @return null|array<int, array<mixed>> The array created from the csv string
     * @TODO throw exception instead of returning null.
     */
    public function excelCsvToArray(string $csvString): ?array
    {
        $csvString = trim($csvString);
        if (!$csvString) {
            $this->setError(__('Invalid or empty CSV String'));

            return null;
        }
        $csvArray = [];

        $rows = explode("\n", $csvString);
        foreach ($rows as $row) {
            $row = trim($row);
            if (!$row) {
                continue;
            }

            $row = str_getcsv($row);

            $emptyCnt = 0;
            foreach ($row as $k => $v) {
                $v = trim($v);
                if (!$v) {
                    $emptyCnt++;
                }
                $row[$k] = $v;
            }
            if ($emptyCnt >= count($row)) {
                continue;
            }
            $csvArray[] = $row;
        }

        return $csvArray;
    }

    /**
     * Converts an excel file into an html table.
     *
     * @param string $inputFileName The path to the excel file
     * @return null|string The generated html
     * @TODO throw exception instead of returning null.
     */
    public function excelFileToHtml(string $inputFileName): ?string
    {
        if (!$inputFileName) {
            $this->setError(__('Unknown File Path'));

            return null;
        }
        if (!is_readable($inputFileName)) {
            $this->setError(__('Unable to read the File'));

            return null;
        }

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);

        $spreadsheet = $reader->load($inputFileName);

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($spreadsheet);
        $writer->setSheetIndex(0);

        ob_start();
        $writer->save('php://output');
        $results = ob_get_contents();
        ob_end_clean();

        return is_string($results) ? $results : null;
    }

    /**
     * Converts and excel file into an array.
     *
     * @param string $inputFileName The path to the excel file.
     * @param bool $includeHiddenRows If we should include the hidden rows in the excel file.
     * @return array<int, array<int|string, string>> The generated array from the excel file.
     * @TODO throw exception instead of returning null.
     */
    public function excelFileToArray(string $inputFileName, bool $includeHiddenRows = false): array
    {
        if (!$inputFileName) {
            $this->setError(__('Unknown File Path'));

            return [];
        }
        if (!is_readable($inputFileName)) {
            $this->setError(__('Unable to read the File'));

            return [];
        }

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);

        $spreadsheet = $reader->load($inputFileName);

        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->getAutoFilter()->showHideRows();

        // go through the active worksheet (which is the first sheet, 0 index by default)
        // retrieve the header names

        $headers = [];
        $alldata = [];

        $first = true;
        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);

            // only work with visible rows
            if (!$includeHiddenRows) {
                if (!$worksheet->getRowDimension($row->getRowIndex())->getVisible()) {
                    continue;
                }
            }

            $rowdata = [];
            // find the column names
            $cell_i = 0;
            foreach ($cellIterator as $cell) {
                $cell_value = $cell->getFormattedValue();

                if ($first) {
                    $headers[] = strtolower(Text::slug($cell_value));
                } else {
                    $cell_key = ($headers[$cell_i] ?? $cell_i);

                    if ($cell_key == 'date') {
                        $cell_value = $this->excelFixDate($cell->getValue());
                    }

                    if ($cell_key) {
                        $rowdata[$cell_key] = $cell_value;
                    }
                }
                $cell_i++;
            }
            if ($rowdata) {
                $alldata[] = $rowdata;
            }

            $first = false;
        }

        return $alldata;
    }

    /**
     * Fixes the date from excel to a proper timestamp.
     *
     * @param string $date The excel date
     * @return string The fixed date for excel
     */
    public function excelFixDate(string $date): string
    {
        $date = trim($date);
        if ($date) {
            $date = Date::excelToTimestamp(floatval($date));
            $date = strtotime('+1 day', $date);
            $date = date('Y-m-d H:i:s', $date);
        }

        return $date;
    }
}
