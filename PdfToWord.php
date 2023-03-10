<?php
include("vendor/autoload.php");

use \ConvertApi\ConvertApi;
session_start();

if(!isset($_SESSION['admin_name'])){
   header('location:login_form.php');
}


ConvertApi::setApiSecret('ENTER_HERE_YOUR_SECRET_KEY');
$ClientSecret="8b2ad2ac989ebd1dccbdb007f9de9555";
$ClientID="27ebdc20-93b7-4b3d-9dde-ad974df50b0e";
$conn = mysqli_connect("localhost", "root", "", "pdfdata");
$mysqli = new mysqli('localhost', 'root', '','pdfdata' );

// Check for connection errors
if ($mysqli->connect_errno) {
    echo "Failed to connect to MySQL: " . $mysqli->connect_error;
    exit();
}

$msg = "";
$contents = "";
$output = "";
if (isset($_POST["submit"])) {
  $filename = $_FILES["formFile"]["name"];
  $filetype = $_FILES["formFile"]["type"];
  $pdf_size=$_FILES['formFile']['size'];
  $filetemp = $_FILES["formFile"]["tmp_name"];
  $dir = 'uploads/' . $filename;

       

  if ($filetype == "application/pdf") {
    move_uploaded_file($filetemp, $dir);
    $sql="INSERT INTO pdf_file(pdf) values('$filename')";
    $query=mysqli_query($mysqli,$sql);
    
    $wordsApi = new Aspose\Words\WordsApi($ClientID,$ClientSecret);
    
    
    $format = "docx";
    $file = $dir;
    
    $request = new Aspose\Words\Model\Requests\ConvertDocumentRequest($file, $format,null);
    $result = $wordsApi->ConvertDocument($request); 
    copy($result->getPathName(),"C:/xampp/htdocs/demo/converted_files/output.docx");
    $fileDoc = 'C:/xampp/htdocs/demo/converted_files/output.docx';

    
/////////////////////////////// 

//MYSQL END/////////////////////////////////
     $msg = "<div class='alert alert-success'>File converted.</div>";
     header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
     header('Content-Disposition: inline; filename="' . basename($fileDoc) . '"');
     header('Content-Length: ' . filesize($fileDoc));
     readfile($fileDoc);

   } 
   else {
     $msg = "<div class='alert alert-danger'>Something wrong.</div>";
   }
      
// Set the MIME type


// Output the file to the browser
    
}

  
  else {
    $msg = "<div class='alert alert-danger'>PLEASE UPLOAD A PDF FILE.</div>";
  }
  


?>
<!doctype html>
<html lang="en">

<head>
  <!-- Required meta tags -->
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">


  
  <link rel="stylesheet" href="css/style.css">

  <title>PDF to WORD - Pure Coding</title>
</head>


<body>
               <div class="container">
               <div class="content">
               <h3>hi, <span>admin</span></h3>
               <h1>welcome <span><?php echo $_SESSION['admin_name'] ?></span></h1>
               <h3>PDF TO WORD </h3>
               <a href="logout.php" class="btnlogout">logout</a>
            <br></br>
            <br></br>
            

            <?php echo "<p style='color:red;'>$msg.</p>"; ?>
            
            <br></br>
            <form action="" method="POST" enctype="multipart/form-data">
              <div class="mb-3">
              <h3>BROWSE FILE </h3>
                <br></br>
                <input class="form-control" type="file" id="formFile" name="formFile" required>
              </div>
              <br></br>
              <button class="btn btn-primary" name="submit">Convert Now</button>
            </form>
           
            <img src="<?php echo $output; ?>" alt="" class="img-fluid">
          </div>
        </div>
      

 
            

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
</body>

</html>
