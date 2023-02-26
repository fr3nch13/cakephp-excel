<?php

declare(strict_types=1);

/**
 * @var \Fr3nch13\Excel\View\ExcelView $this
 */

$properties = $properties ?? [];
$headers = $headers ?? [];
$rows = $rows ?? [];

$this->Excel->create($properties, $headers, $rows)
    ->download('Xlsx');
