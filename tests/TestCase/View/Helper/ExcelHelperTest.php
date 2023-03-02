<?php
declare(strict_types=1);

/**
 * ExcelHelperTest
 */

namespace Fr3nch13\Excel\Test\TestCase\View\Helper;

use Cake\Http\ServerRequest;
use Cake\Routing\Router;
use Cake\View\View;
use Fr3nch13\Excel\View\Helper\ExcelHelper;

/**
 * ExcelHelper Test Class
 *
 * This class contains the main tests for the ExcelHelper Class.
 */
class ExcelHelperTest extends \Cake\TestSuite\TestCase
{
    /**
     * @var \Cake\View\View
     */
    public $View;

    /**
     * @var \Fr3nch13\Excel\View\Helper\ExcelHelper
     */
    public $ExcelHelper;

    /**
     * Setup the application so that we can run the tests.
     *
     * The setup involves initializing a new CakePHP view and using that to
     * get a copy of the ExcelHelper.
     */
    public function setUp(): void
    {
        parent::setUp();

        Router::reload();
        $request = new ServerRequest();
        Router::setRequest($request);

        $this->View = new View($request);
        $this->ExcelHelper = new ExcelHelper($this->View);

        static::setAppNamespace();
        $this->loadPlugins(['Fr3nch13/Excel' => []]);
    }

    /**
     * tearDown method
     *
     * @return void
     */
    public function tearDown(): void
    {
        parent::tearDown();

        $this->clearPlugins();
        unset($this->ExcelHelper, $this->View);
    }

    /**
     * ['controller' => 'background'] in the options is needed for testing with a mapable route.
     *
     * @return void
     */
    public function testCreate(): void
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

        $excel = $this->ExcelHelper->create($properties, $headers, $rows);

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
        $this->assertSame(['sheet_data' => $sheet_data], $excel->toArray(null));
        $this->assertSame([], $excel->toArray(1));
    }
}
