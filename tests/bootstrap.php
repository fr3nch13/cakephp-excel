<?php
declare(strict_types=1);

/**
 * Test suite bootstrap.
 *
 * These are the specific settings for this plugin.
 * This uses fr3nch13/cakephp-pta to provide a generic application for testing.
 * Setting passed to cakephp-pta's bootstrap and application are defined here.
 */

use Cake\Core\Configure;

// Configure your stuff here for the plugin_bootstrap.php below.
define('TESTS', __DIR__ . DS);

Configure::write('Tests.Plugins', [
    'Fr3nch13/Excel',
]);

Configure::write('Tests.Helpers', [
    'Excel' => ['className' => 'Fr3nch13/Excel.Excel'],
]);

////// Ensure we can setup an environment for the Test Application instance.
$root = dirname(__DIR__);
chdir($root);
require_once $root . '/vendor/fr3nch13/cakephp-pta/tests/plugin_bootstrap.php';
