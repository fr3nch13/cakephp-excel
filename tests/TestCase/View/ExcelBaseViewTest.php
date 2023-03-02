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
        $request = new ServerRequest();
        Router::setRequest($request);

        $this->View = new ExcelBaseView($request);

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
        unset($this->View);
    }

    /**
     * ['controller' => 'background'] in the options is needed for testing with a mapable route.
     *
     * @return void
     */
    public function testView(): void
    {
        $this->assertInstanceOf(\Fr3nch13\Excel\View\Helper\ExcelHelper::class, $this->View->Excel);
    }
}
