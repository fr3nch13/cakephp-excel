<?php
declare(strict_types=1);

/**
 * PluginTest
 */

namespace Fr3nch13\Excel\Test\TestCase;

use Cake\Console\ConsoleIo;
use Cake\Console\TestSuite\StubConsoleOutput;
use Cake\I18n\FrozenTime;
use Cake\TestSuite\IntegrationTestTrait;
use Cake\TestSuite\TestCase;
use Fr3nch13\Excel\Excel;
use Fr3nch13\Excel\Exception\ExcelException;

/**
 * PluginTest class
 */
class ExcelTest extends TestCase
{
    /**
     * Apparently this is the new Cake way to do integration.
     */
    use IntegrationTestTrait;

    /**
     * @var \Cake\Console\TestSuite\StubConsoleOutput
     */
    protected $stub;

    /**
     * @var \Cake\Console\ConsoleIo
     */
    protected $io;

    /**
     * @var string Path to the valid excel file.
     */
    public $filepath;

    /**
     * @var \Fr3nch13\Excel\Excel A valid excel object
     */
    public $Excel;

    /**
     * setUp method
     *
     * @return void
     */
    public function setUp(): void
    {
        parent::setUp();

        $this->stub = new StubConsoleOutput();
        $this->io = new ConsoleIo($this->stub);
        $this->io->level(ConsoleIo::VERBOSE);

        $this->filepath = dirname(__DIR__) . DS . 'assets' . DS . 'FinancialSample.xlsx';
    }

    /**
     * tearDown method
     *
     * @return void
     */
    public function tearDown(): void
    {
        parent::tearDown();
    }

    /**
     * testBootstrap
     *
     * @return void
     */
    public function testConstructor(): void
    {
        $excel = new Excel();

        $this->assertSame('', $excel->getFilePath());
        $this->assertSame(0, $excel->getSheetIndex());
        $this->assertNull($excel->getLineLimit());
        $this->assertSame([], $excel->getErrors());

        $excel = new Excel($this->filepath);

        $this->assertSame($this->filepath, $excel->getFilePath());
        $this->assertSame(0, $excel->getSheetIndex());
        $this->assertNull($excel->getLineLimit());
        $this->assertSame([], $excel->getErrors());

        $excel = new Excel($this->filepath, 1);

        $this->assertSame($this->filepath, $excel->getFilePath());
        $this->assertSame(1, $excel->getSheetIndex());
        $this->assertNull($excel->getLineLimit());

        $excel = new Excel($this->filepath, 2);

        $this->assertSame($this->filepath, $excel->getFilePath());
        $this->assertSame(2, $excel->getSheetIndex());
        $this->assertNull($excel->getLineLimit());

        $excel = new Excel($this->filepath, 2, 100);

        $this->assertSame($this->filepath, $excel->getFilePath());
        $this->assertSame(2, $excel->getSheetIndex());
        $this->assertSame(100, $excel->getLineLimit());

        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Error unable to read file: `');
        $excel = new Excel('/does/not/exist');
    }

    /**
     * Test loading a spreadsheet
     *
     * @return void
     */
    public function testLoadNoFile(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Error unable to read file: `');
        $excel = new Excel('/does/not/exist');
        $excel->load();
    }

    /**
     * Test loading a spreadsheet
     *
     * @return void
     */
    public function testLoadBadFile(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Unable to identify a reader for this file');
        $excel = new Excel(__FILE__);
        $this->assertInstanceOf(Excel::class, $excel->load());
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Reader\Xlsx::class, $excel->getReader());
    }

    /**
     * Test loading a spreadsheet
     *
     * @return void
     */
    public function testLoadBadFilePhp(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Unable to identify a reader for this file');
        $excel = new Excel(dirname(__DIR__) . DS . 'bootstrap.php');
        $this->assertInstanceOf(Excel::class, $excel->load());
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Reader\Xlsx::class, $excel->getReader());
    }

    /**
     * Test loading a spreadsheet
     *
     * @return void
     */
    public function testLoad(): void
    {
        $excel = new Excel($this->filepath);
        $this->assertInstanceOf(Excel::class, $excel->load());
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Reader\Xlsx::class, $excel->getReader());
    }

    /**
     * Test loading a spreadsheet
     *
     * @return void
     */
    public function testLoadBadSheetIndex(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('You tried to set a sheet active by the out of bounds index: 1. ' .
            'The actual number of sheets is 1.');
        $excel = new Excel($this->filepath, 1);
        $excel->load();
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testCreateEmpty(): void
    {
        $excel = new Excel();
        $excel->create();
        $properties = $excel->getProperties();
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Document\Properties::class, $properties);
        $this->assertSame('', $excel->getFilePath());
        $this->assertSame('', $excel->getProperties('category'));
        $this->assertSame('', $excel->getProperties('company'));
        $this->assertGreaterThan(1677624438, $excel->getProperties('created'));
        $this->assertSame('Unknown Creator', $excel->getProperties('creator'));
        $this->assertSame('', $excel->getProperties('description'));
        $this->assertSame([], $excel->getProperties('keywords'));
        $this->assertSame('Unknown Creator', $excel->getProperties('last_modified_by'));
        $this->assertSame('Unknown Creator', $excel->getProperties('modifier'));
        $this->assertSame('', $excel->getProperties('subject'));
        $this->assertSame('Untitled Spreadsheet', $excel->getProperties('title'));
        $this->assertSame('', $excel->getProperties('custom'));

        $sheet_data = [[]];
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray());
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray(0));
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray(null, $this->io));
        $this->assertSame([], $excel->toArray(1));
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testCreateWithDetails(): void
    {
        $properties = [
            'category' => 'Testing Category',
            'company' => 'Fr3nch13 Inc',
            'created' => 1677624438,
            'creator' => 'fr3nch13 createor',
            'description' => 'this is a test',
            'keywords' => ['cakephp', 'test'],
            'modifier' => 'fr3nch13 modifier',
            'subject' => 'Testing Subject',
            'title' => 'Test Title',
            'custom key' => 'custom value',
            'blahblah' => 'blahblah value',
            'sheet_title' => 'Test Sheet Title',
        ];
        $headers = [
            'column1' => 'Header 1',
            'column2' => 'Header 2',
        ];
        $rows = [
            ['column1' => 'c1r1', 'column2' => 'c2r1'],
            ['column1' => 'c1r2', 'column2' => 'c2r2'],
        ];

        $excel = new Excel();
        $excel->create($properties, $headers, $rows);
        $properties = $excel->getProperties();
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Document\Properties::class, $properties);
        $this->assertSame('', $excel->getFilePath());
        $this->assertSame('Testing Category', $excel->getProperties('category'));
        $this->assertSame('Fr3nch13 Inc', $excel->getProperties('company'));
        $this->assertSame(1677624438, $excel->getProperties('created'));
        $this->assertSame('fr3nch13 createor', $excel->getProperties('creator'));
        $this->assertSame('this is a test', $excel->getProperties('description'));
        $this->assertSame(['cakephp', 'test'], $excel->getProperties('keywords'));
        $this->assertSame('fr3nch13 modifier', $excel->getProperties('last_modified_by'));
        $this->assertSame('fr3nch13 modifier', $excel->getProperties('modifier'));
        $this->assertSame('Testing Subject', $excel->getProperties('subject'));
        $this->assertSame('Test Title', $excel->getProperties('title'));
        $this->assertSame('blahblah value', $excel->getProperties('blahblah'));
        $this->assertSame('custom value', $excel->getProperties('custom key'));

        $sheet_data = [
            [
                1 => ['A' => 'Header 1', 'B' => 'Header 2'],
                2 => ['A' => 'c1r1', 'B' => 'c2r1'],
                3 => ['A' => 'c1r2', 'B' => 'c2r2'],
            ],
        ];
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray());
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray(0));
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray(null, $this->io));
        $this->assertSame([], $excel->toArray(1));
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testCreateWithBadColumns(): void
    {
        $headers = ['column1' => 'Header 1', 'column2' => 'Header 2'];
        $rows = [
            ['column1' => 'c1r1', 'column2' => 'c2r1', 'badcolumn' => 'c3r1'],
        ];

        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Cells longer than header row. Row:0');
        $excel = new Excel();
        $excel->create([], $headers, $rows);
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testSaveEmpty(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Error when loading excel file: Error: `File "" does not exist.`. File: ``');
        $excel = new Excel();
        $excel->save();
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testSaveNoPath(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Unable to save the spreadsheet. Set the filePath first.');
        $excel = $this->makeSampleExcel();
        $excel->save();
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testSaveBadPath(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Unable to write to: `');

        $filepath = dirname($this->filepath) . DS . 'dontexist' . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testSaveBadType(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Unknown type: `Txt`');

        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath, 'Txt');
    }

    /**
     * Test creating a spreadsheet
     *
     * @return void
     */
    public function testSave(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));
    }

    /**
     * Test creating a spreadsheet.
     * Saved a file like how we did above to ensure we can read to array.
     *
     * @return void
     */
    public function testExcelFileToArray(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $excel = new Excel();
        $results = $excel->excelFileToArray($filepath);

        $expected = [
            ['header-1' => 'c1r1', 'header-2' => 'c2r1'],
            ['header-1' => 'c1r2', 'header-2' => 'c2r2'],
        ];
        $this->assertSame($expected, $results);

        $excel = new Excel();
        $results = $excel->excelFileToArray($filepath, true);

        $expected = [
            ['header-1' => 'c1r1', 'header-2' => 'c2r1'],
            ['header-1' => 'c1r2', 'header-2' => 'c2r2'],
        ];
        $this->assertSame($expected, $results);
    }

    /**
     * Test creating a spreadsheet.
     * Saved a file like how we did above to ensure we can read to array.
     *
     * @return void
     */
    public function testDownloadDefault(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $this->expectOutputRegex('#xl/_rels/workbook.xml.rels#');
        $excel->download();
    }

    /**
     * Test downloading a XLSX file.
     * Saved a file like how we did above to ensure we can read to array.
     *
     * @return void
     */
    public function testDownloadExcel(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $this->expectOutputRegex('#xl/_rels/workbook.xml.rels#');
        $excel->download('Xlsx');
    }

    /**
     * Test downloading a PDF file.
     * Saved a file like how we did above to ensure we can read to array.
     *
     * @return void
     */
    public function testDownloadPdf(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $this->expectOutputRegex('#\%PDF\-1\.4#');
        $excel->download('Pdf');
    }

    /**
     * Test downloading a CSV file.
     * Saved a file like how we did above to ensure we can read to array.
     *
     * @return void
     */
    public function testDownloadCsv(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $this->expectOutputRegex('#\"Header 1\"\,\"Header 2\"#');
        $excel->download('Csv');
    }

    /**
     * Test updating an excel file.
     *
     * @return void
     */
    public function testUpdateFromArray(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.xlsx';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath);
        $this->assertTrue(file_exists($filepath));

        $excel = new Excel($filepath);
        $results = $excel->toArray();

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('Header 1', $results['sheet_data'][0][1]['A']);
        $this->assertSame('c1r1', $results['sheet_data'][0][2]['A']);

        // update a single column
        $updateRows = [
            2 => ['A' => 'c1r1 updated'],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('Header 1', $results['sheet_data'][0][1]['A']);
        $this->assertSame('c1r1 updated', $results['sheet_data'][0][2]['A']);

        // add a header
        $updateRows = [
            1 => ['C' => 'Timestamp'],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('Timestamp', $results['sheet_data'][0][1]['C']);

        // updating rows with a timestamp
        $updateRows = [
            2 => ['C' => new FrozenTime('01/01/2023')],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('2023/01/01', $results['sheet_data'][0][2]['C']);

        // updating rows with a timestamp
        $updateRows = [
            2 => ['C' => new FrozenTime('01/01/1900')],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('1900/01/01', $results['sheet_data'][0][2]['C']);

        // updating rows with a array info
        $updateRows = [
            2 => [
                'A' => ['value' => 'c1r2 updated', 'options' => ['lock' => false, 'hidden' => true]],
                'B' => ['value' => 'c2r2 updated', 'options' => ['lock' => true, 'hidden' => false]],
            ],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('c1r2 updated', $results['sheet_data'][0][2]['A']);

        // updating rows with a array info
        $updateRows = [
            2 => [
                'A' => ['value' => 'c1r2 updated', 'options' => ['lock' => false, 'hidden' => true]],
                'B' => ['value' => 'c2r2 updated', 'options' => ['lock' => true, 'hidden' => false]],
            ],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('c1r2 updated', $results['sheet_data'][0][2]['A']);

        $excel = new Excel();

        $results = $excel->excelFileToArray($filepath);

        $this->assertSame(2, count($results));
        $this->assertSame('c1r2 updated', $results[0]['header-1']);

        // Test Non int/string values
        $updateRows = [
            2 => [
                'A' => false,
                'B' => null,
            ],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('FALSE', $results['sheet_data'][0][2]['A']);
        $this->assertSame(null, $results['sheet_data'][0][2]['B']);

        // Test Bad Value
        $updateRows = [
            2 => [
                'A' => [],
            ],
        ];

        $excel = new Excel($filepath);
        $excel->updateFromArray($updateRows, $this->io);
        $results = $excel->toArray();
        $excel->save($filepath);

        $this->assertSame(3, count($results['sheet_data'][0]));
        $this->assertSame('FALSE', $results['sheet_data'][0][2]['A']);
        $this->assertSame(null, $results['sheet_data'][0][2]['B']);
    }

    /**
     * Testing CSV to Array.
     *
     * @return void
     */
    public function testExcelCSVToArray(): void
    {
        $filepath = dirname($this->filepath) . DS . 'testsave.csv';
        $excel = $this->makeSampleExcel();
        $excel->save($filepath, 'Csv');
        $this->assertTrue(file_exists($filepath));

        $results = null;
        $content = file_get_contents($filepath);
        if ($content) {
            $results = (new Excel())->excelCsvToArray($content);
        }

        $expected = [
            ['Header 1', 'Header 2'],
            ['c1r1', 'c2r1'],
            ['c1r2', 'c2r2'],
        ];
        $this->assertSame($expected, $results);

        $results = null;
        $filepath = dirname($this->filepath) . DS . 'bad.csv';
        $content = file_get_contents($filepath);
        if ($content) {
            $results = (new Excel())->excelCsvToArray($content);
        }

        $expected = [
            ['Header 1', 'Header 2'],
            ['c1r1', 'c2r1'],
            ['c1r2', 'c2r2'],
        ];
        $this->assertSame($expected, $results);

        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Invalid or empty CSV String');
        $results = (new Excel())->excelCsvToArray('');

        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Invalid or empty CSV String');
        $results = (new Excel())->excelCsvToArray(' ');
    }

    /**
     * Testing line limit.
     *
     * @return void
     */
    public function testSetLineLimits(): void
    {
        $excel = new Excel($this->filepath, 0, 10);
        $results = $excel->toArray();

        $this->assertSame(10, count($results['sheet_data'][0]));
    }

    /**
     * Test all of the helper methods.
     * Including the getters/setters.
     */

    /**
     * Test Memory
     *
     * @return void
     */
    public function testMemory(): void
    {
        $excel = new Excel();
        $this->assertGreaterThan(100, $excel->memoryUsage(false));
        $this->assertSame('1 KB', $excel->memoryUsage(true, 1024));
        $this->assertSame('1 MB', $excel->memoryUsage(true, 1024 * 1024));
        $this->assertSame('1 GB', $excel->memoryUsage(true, 1024 * 1024 * 1024));
    }

    /**
     * Test Memory
     *
     * @return void
     */
    public function testSheetNamesNoFile(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Error when loading excel file: Error: `File "" does not exist.`. File: ``');
        $excel = new Excel();
        $excel->getSheetNames();
    }

    /**
     * Test Memory
     *
     * @return void
     */
    public function testSheetNames(): void
    {
        $excel = new Excel($this->filepath);
        $names = $excel->getSheetNames();

        $this->assertSame(['Sheet1'], $names);
    }

    /**
     * Test Memory
     *
     * @return void
     */
    public function testGetReaderNoFile(): void
    {
        $this->expectException(ExcelException::class);
        $this->expectExceptionCode(500);
        $this->expectExceptionMessage('Error when loading excel file: Error: `File "" does not exist.`. File: ``');
        $excel = new Excel();
        $excel->getReader();
    }

    /**
     * Test Memory
     *
     * @return void
     */
    public function testGetReader(): void
    {
        $excel = new Excel($this->filepath);
        $reader = $excel->getReader();
        $this->assertInstanceOf(\PhpOffice\PhpSpreadsheet\Reader\Xlsx::class, $reader);
    }

    /**
     * Test Memory
     *
     * @return void
     */
    public function testLineLimit(): void
    {
        $excel = new Excel();

        $this->assertInstanceOf(Excel::class, $excel->setLineLimit(1000));
        $this->assertSame(1000, $excel->getLineLimit());
    }

    /**
     * Test File Path
     *
     * @return void
     */
    public function testFilePath(): void
    {
        $excel = new Excel();

        $this->assertInstanceOf(Excel::class, $excel->setFilePath($this->filepath));
        $this->assertSame($this->filepath, $excel->getFilePath());
    }

    /**
     * Test Sheet Index
     *
     * @return void
     */
    public function testSeetIndex(): void
    {
        $excel = new Excel();

        $this->assertInstanceOf(Excel::class, $excel->setSheetIndex(1));
        $this->assertSame(1, $excel->getSheetIndex());
    }

    /**
     * Test Error
     *
     * @return void
     */
    public function testError(): void
    {
        $excel = new Excel();

        $this->assertInstanceOf(Excel::class, $excel->setError('Error Message'));
        $this->assertInstanceOf(Excel::class, $excel->setError('Error Message 2'));
        $this->assertSame([
            'Error Message',
            'Error Message 2',
        ], $excel->getErrors());
    }

    /**
     * Test Error
     *
     * @return void
     */
    public function testExcelFixDate(): void
    {
        $excel = new Excel();

        $this->assertSame('2022-08-17 00:00:00', $excel->excelFixDate('44790'));
        $this->assertSame('2022-08-17 00:00:00', $excel->excelFixDate(44790));

        $this->assertSame('2014-01-01 00:00:00', $excel->excelFixDate('41640'));
        $this->assertSame('2014-01-01 00:00:00', $excel->excelFixDate(41640));
    }

    /**
     * Creates a sample excel file to test against.
     *
     * @return \Fr3nch13\Excel\Excel
     */
    public function makeSampleExcel(): Excel
    {
        $properties = [
            'category' => 'Testing Category',
            'company' => 'Fr3nch13 Inc',
            'created' => 1677624438,
            'creator' => 'fr3nch13 createor',
            'description' => 'this is a test',
            'keywords' => ['cakephp', 'test'],
            'modifier' => 'fr3nch13 modifier',
            'subject' => 'Testing Subject',
            'title' => 'Test Title',
            'custom key' => 'custom value',
            'blahblah' => 'blahblah value',
            'sheet_title' => 'Test Sheet Title',
        ];
        $headers = [
            'column1' => 'Header 1',
            'column2' => 'Header 2',
        ];
        $rows = [
            ['column1' => 'c1r1', 'column2' => 'c2r1'],
            ['column1' => 'c1r2', 'column2' => 'c2r2'],
        ];

        $excel = new Excel();

        return $excel->create($properties, $headers, $rows);
    }
}
