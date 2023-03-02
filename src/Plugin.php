<?php
declare(strict_types=1);

/**
 * Plugin Definitions
 */

namespace Fr3nch13\Excel;

use Cake\Core\BasePlugin;
use Cake\Core\Configure;

/**
 * Plugin Definitions
 */
class Plugin extends BasePlugin
{
    /**
     * Bootstraping for this specific plugin.
     *
     * @param \Cake\Core\PluginApplicationInterface $app The app object.
     * @return void
     */
    public function bootstrap(\Cake\Core\PluginApplicationInterface $app): void
    {
        // By default will load `config/bootstrap.php` in the plugin.
        parent::bootstrap($app);

        // Add constants, load configuration defaults.
        if (!Configure::read('Excel')) {
            Configure::write('Excel', [
                'test' => 'TEST',
                'modifiedBy' => 'Fr3nch13',
            ]);
        }
    }
}
