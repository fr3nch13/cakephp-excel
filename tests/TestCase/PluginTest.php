<?php
declare(strict_types=1);

/**
 * PluginTest
 */

namespace Fr3nch13\Excel\Test\TestCase;

use App\Application;
use Cake\Core\Configure;
use Cake\TestSuite\IntegrationTestTrait;
use Cake\TestSuite\TestCase;

/**
 * PluginTest class
 */
class PluginTest extends TestCase
{
    /**
     * Apparently this is the new Cake way to do integration.
     */
    use IntegrationTestTrait;

    /**
     * setUp method
     *
     * @return void
     */
    public function setUp(): void
    {
        parent::setUp();
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
    public function testBootstrap(): void
    {
        $app = new Application(CONFIG);
        $app->bootstrap();
        $app->pluginBootstrap();
        $plugins = $app->getPlugins();

        $this->assertSame('Fr3nch13/Excel', $plugins->get('Fr3nch13/Excel')->getName());

        // make sure it was able to read and store the config.
        $this->assertEquals(Configure::read('Excel.test'), 'TEST');
    }
}
