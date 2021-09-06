<?php 
namespace App\Models;
use CodeIgniter\Database\ConnectionInterface;
use CodeIgniter\Model;
 
class StudentModel extends Model
{
    protected $table = 'Students';
 
    protected $allowedFields = ['first_name','last_name','address','email', 'mobile'];
}
?>