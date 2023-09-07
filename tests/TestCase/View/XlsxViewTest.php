<?php
declare(strict_types=1);

/**
 * XlsxViewTest
 */

namespace Fr3nch13\Excel\Test\TestCase\View\Helper;

use Fr3nch13\Excel\View\XlsxView;

/**
 * XlsxView Test Class
 *
 * This class contains the main tests for the XlsxView Class.
 */
class XlsxViewTest extends \Cake\TestSuite\TestCase
{
    /**
     * @var \Fr3nch13\Excel\View\XlsxView
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

        $this->View = new XlsxView(null, null, null, $viewOptions);
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
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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

        // should use our layout and the app's template.
        $this->assertSame(
            '/templates/layout/xlsx/default.php',
            str_replace(PLUGIN_ROOT, '', $filenames['layout'])
        );

        // since there are no routes, the Request can't route to a controller.
        // here, we're ensuring the subdir is set.
        $this->assertSame(
            '/templates/Tests/xlsx/index.php',
            str_replace(PLUGIN_ROOT, '', $filenames['template'])
        );
    }
}
