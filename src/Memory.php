<?php
declare(strict_types=1);

/**
 * MemoryBehavior
 */

namespace Fr3nch13\Excel;

/**
 * Tracks Memory and time usage, mainly for cron jobs.
 */
class Memory
{
    /**
     * Default config settings.
     *
     * @var array<string, mixed>
     */
    protected $_defaultConfig = [];

    /**
     * Trackes the start times.
     * Format is ['(time_key)' => (unix timestamp)].
     *
     * @var array<string, int>
     */
    protected $_startTimes = [];

    /**
     * Trackes the end times.
     * Format is ['(time_key)' => (unix timestamp)].
     *
     * @var array<string, int>
     */
    protected $_endTimes = [];

    /**
     * Tracks the highest recorded memory usage from when memoryUsage was called.
     *
     * @var float
     */
    protected $memUsageHighest = 0;

    /**
     * Reports the memory usage at the time it is called.
     *
     * @param bool $nice If we should return the bytes (false), of the calculated amount in a nice format (true).
     * @param float|null $mem_usage The memory number to be made nice.
     * @return string the memory usage stat.
     */
    public function memoryUsage(bool $nice = true, ?float $mem_usage = null): string
    {
        if (!$mem_usage) {
            $mem_usage = memory_get_usage();
        }
        // track the highest usage.
        if ($this->memUsageHighest < $mem_usage) {
            $this->memUsageHighest = $mem_usage;
        }
        if ($nice) {
            if ($mem_usage < 1024) {
                $mem_usage = $mem_usage . ' B';
            } elseif ($mem_usage < 1048576) {
                $mem_usage = round($mem_usage / 1024, 2) . ' KB';
            } elseif ($mem_usage < 1073741824) {
                $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
            } else {
                $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
            }
        }

        return strval($mem_usage);
    }

    /**
     * Reports the highest memory usage.
     *
     * @param bool $nice If we should return the bytes (false), of the calculated amount in a nice format (true).
     * @return string the highest memory usage stat.
     */
    public function memoryUsageHighest($nice = true): string
    {
        return $this->memoryUsage($nice, $this->memUsageHighest);
    }

    /**
     * Reports the peak memory usage at the time it is called.
     *
     * @param bool $nice If we should return the bytes (false), of the calculated amount in a nice format (true).
     * @param null|int $mem_usage The memory number to be made nice.
     * @return string the slugged string
     */
    public function memoryPeakUsage($nice = true, $mem_usage = null): string
    {
        if (!$mem_usage) {
            $mem_usage = memory_get_usage(true);
        }
        if ($nice) {
            if ($mem_usage < 1024) {
                $mem_usage = $mem_usage . ' B';
            } elseif ($mem_usage < 1048576) {
                $mem_usage = round($mem_usage / 1024, 2) . ' KB';
            } elseif ($mem_usage < 1073741824) {
                $mem_usage = round($mem_usage / 1048576, 2) . ' MB';
            } else {
                $mem_usage = round($mem_usage / 1073741824, 2) . ' GB';
            }
        }

        return strval($mem_usage);
    }
}
