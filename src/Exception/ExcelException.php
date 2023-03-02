<?php
declare(strict_types=1);

/**
 * Excel Exception
 */

namespace Fr3nch13\Excel\Exception;

use Cake\Core\Exception\CakeException;

/**
 * Excel Exception
 *
 * Used for tracking problems specific to this plugin.
 * This uses CakeException so it can be handled, and seen/hidden via CakePHP.
 */
class ExcelException extends CakeException
{
    /**
     * Default exception code
     *
     * @var int
     */
    protected $_defaultCode = 500;
}
