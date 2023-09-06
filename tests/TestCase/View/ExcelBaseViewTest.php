<?php
declare(strict_types=1);

/**
 * ExcelBaseViewTest
 */

namespace Fr3nch13\Excel\Test\TestCase\View\Helper;

use Cake\Http\ServerRequest;
use Cake\Routing\Router;
use Fr3nch13\Excel\View\ExcelBaseView;

/**
 * ExcelBaseView Test Class
 *
 * This class contains the main tests for the ExcelBaseView Class.
 */
class ExcelBaseViewTest extends \Cake\TestSuite\TestCase
{
    /**
     * @var \Fr3nch13\Excel\View\ExcelBaseView
     */
    public $View;

    /**
     * Setup the application so that we can run the tests.
     *
     * The setup involves initializing a new CakePHP view and using that to
     * get a copy of the ExcelBaseView.
     */
    public function setUp(): void
    {
        parent::setUp();

        Router::reload();
        $request = new ServerRequest([
            'url' => '/pages/home',
        ]);
        Router::setRequest($request);

        $this->View = new ExcelBaseView($request);

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
            'text/html',
            $this->View->getResponse()->getType()
        );
    }

    /**
     * Test The Layour file.
     *
     * @return void
     */
    public function testLayout(): void
    {
        $this->assertSame(
            'Fr3nch13/Excel.default',
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
        $this->View->setTemplate('base');
        $this->View->render();

        $filenames = $this->View->getFileNames();

        // These should be using the App's layout and template.
        $this->assertSame(
            '/vendor/fr3nch13/cakephp-pta/tests/test_app/templates/layout/default.php',
            str_replace(PLUGIN_ROOT, '', $filenames['layout'])
        );
        $this->assertSame(
            '/vendor/fr3nch13/cakephp-pta/tests/test_app/templates/base.php',
            str_replace(PLUGIN_ROOT, '', $filenames['template'])
        );
    }
}
