<?php


namespace App\Services;


use ReflectionClass;

class Business
{
  private static $yahooStickers =

    public function getAllBusiness(): array
    {
        return (new ReflectionClass(self::class))->getStaticProperties();
    }
}