<?php

namespace App\Exports;

use App\User;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\BeforeSheet;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithHeadings;

class UsersPerMonthSheet implements FromQuery, WithTitle, WithHeadings, WithEvents
{

    /**
     * @return array
     */
    public function registerEvents(): array
    {
        $styleArray = [
            'font' => [
                'bold' => true,
            ]
        ];

        return [
            // Handle by a closure.
            AfterSheet::class => function (AfterSheet $event) use ($styleArray) {
                // $event->sheet->insertNewRowBefore(7, 2);
                // $event->sheet->insertNewColumnBefore('A', 2);
                $event->sheet->getStyle('A1:G1')->applyFromArray($styleArray);
                $event->sheet->setCellValue('E27', '=SUM(E2:E26)');
            },
        ];
    }

     /**
     * @return Builder
     */
    public function query()
    {
        return User
            ::query()
            ->where('id', '>', 25);
    }

    public function headings(): array
    {
        return [
            'id',
            'Name',
            'Email',
            'Verified At',
            'Points',
            'Created At',
            'Updated At',
        ];
    }

    /**
     * @return string
     */
    public function title(): string
    {
        return 'Month';
    }
}
