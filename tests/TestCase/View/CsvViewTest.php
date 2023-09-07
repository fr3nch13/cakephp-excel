<?php
declare(strict_types=1);

/**
 * CsvViewTest
 */

namespace Fr3nch13\Excel\Test\TestCase\View\Helper;

use Cake\Http\ServerRequest;
use Cake\Routing\Router;
use Fr3nch13\Excel\View\CsvView;

/**
 * CsvView Test Class
 *
 * This class contains the main tests for the CsvView Class.
 */
class CsvViewTest extends \Cake\TestSuite\TestCase
{
    /**
     * @var \Fr3nch13\Excel\View\CsvView
     */
    public $View;

    /**
     * Setup the application so that we can run the tests.
     *
     * The setup involves initializing a new CakePHP view and using that to
     * get a copy of the CsvView.
     */
    public function setUp(): void
    {
        parent::setUp();

        Router::reload();
        $request = new ServerRequest();
        Router::setRequest($request);

        $this->View = new CsvView($request);

        static::setAppNamespace();
        $this->loadPlugins(['Fr3nch13/Excel' => []]);
    }

    /**
     * Test that the help is loaded.
     *
     * @return void
     */
    public function testHelper(): void
    {
        $this->assertInstanceOf(
            \Fr3nch13\Excel\View\Helper\ExcelHelper::class,
            $this->View->Excel
        );
    }

    /**
     * Test Content-Type.
     *
     * @return void
     */
    public function testContentType(): void
    {
        $this->assertSame(
            'text/csv',
            $this->View->getResponse()->getType()
        );
    }

    /**
     * Test The Layout file.
     *
     * @return void
     */
    public function testLayout(): void
    {
        $this->assertSame(
            'Fr3nch13/Excel.csv/default',
            $this->View->getLayout()
        );
    }

    /**
     * Test the file paths that are going to be used.
     *
     * @return void
     */
    public function testFileNames(): void
    {
        $this->View->setTemplate('Fr3nch13/Excel.test');
        $this->View->render();

        $filenames = $this->View->getFileNames();

        // should use our layout and the app's template.
        $this->assertSame(
            '/templates/layout/csv/default.php',
            str_replace(PLUGIN_ROOT, '', $filenames['layout'])
        );

        // since there are no routes, the Request can't route to a controller.
        // here, we're ensuring the subdir is set.
        $this->assertSame(
            '/templates/csv/test.php',
            str_replace(PLUGIN_ROOT, '', $filenames['template'])
        );
    }
}
