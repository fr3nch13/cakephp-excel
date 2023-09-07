<?php
declare(strict_types=1);

/**
 * ExcelBaseViewTest
 */

namespace Fr3nch13\Excel\Test\TestCase\View\Helper;

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
     * Setup the View so that we can run the tests.
     */
    public function setUp(): void
    {
        parent::setUp();

        $this->loadPlugins(['Fr3nch13/Excel' => []]);

        $viewOptions = [
            'plugin' => 'Fr3nch13/Excel',
            'name' => 'Tests',
            'templatePath' => 'Tests',
            'template' => 'index',
        ];

        $this->View = new ExcelBaseView(null, null, null, $viewOptions);
    }

    /**
     * Test that the helper is loaded.
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
     * Test The layout.
     *
     * @return void
     */
    public function testLayout(): void
    {
        $this->assertSame(
            'default',
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
        $filenames = $this->View->getFileNames();

        // These should be using the App's layout.
        $this->assertSame(
            '/vendor/fr3nch13/cakephp-pta/tests/test_app/templates/layout/default.php',
            str_replace(PLUGIN_ROOT, '', $filenames['layout'])
        );
        $this->assertSame(
            '/templates/Tests/index.php',
            str_replace(PLUGIN_ROOT, '', $filenames['template'])
        );
    }
}
