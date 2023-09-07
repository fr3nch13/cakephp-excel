<?php

declare(strict_types=1);

/**
 * @var \Fr3nch13\Excel\View\CsvView $this
 */

$properties = [];

if (!$this->fetch('page-title')) {
    $this->assign('page-title', __('Tests'));
}
if (!$this->fetch('page-subtitle')) {
    $this->assign('page-subtitle', __('All'));
}

$properties['title'] = __('{0} - {1}', [
    $this->fetch('page-title'),
    $this->fetch('page-subtitle'),
]);

$rows = [];
for ($i = 0; $i++; $i < 10) {
    $rows[$i] = [
        $i,
        __('Name {0}', [$i]),
        __('Desc {0}', [$i]),
    ];
}

$headers = [
    'id' => __('UID'),
    'name' => __('Name'),
    'slug' => __('Slug'),
    'active' => __('Active'),
    'created' => __('Created'),
];

echo $this->element('Fr3nch13/Excel.csv-index', [
    'options' => [],
    'properties' => $properties,
    'headers' => $headers,
    'rows' => $rows,
]);
