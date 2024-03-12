<?php
require 'vendor/autoload.php'; // Inclure PHPWord

use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\IOFactory;

// Vérifie si le dossier "formulaire" existe, sinon le crée
if (!file_exists('formulaire')) {
    mkdir('formulaire', 0777, true); // Crée le dossier avec les permissions d'accès 0777 (pour permettre la lecture, l'écriture et l'exécution)
}

// Traitement des données du formulaire
$name = isset($_POST['name']) ? $_POST['name'] : '';
$email = isset($_POST['email']) ? $_POST['email'] : '';
$subject = isset($_POST['subject']) ? $_POST['subject'] : '';
$message = isset($_POST['message']) ? $_POST['message'] : '';

// Création du document Word
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$section = $phpWord->addSection();

// Ajout du nom dans le document Word
$section->addText('Nom : ' . $name);

// Ajout des données du formulaire
$section->addText('Email : ' . $email);
$section->addText('Objet : ' . $subject);
$section->addText('Message : ' . $message);

// Nom du fichier Word basé sur le champ "name" du formulaire et la date actuelle
$filename = 'formulaire/' . $name . '_' . date('Y-m-d') . '.docx';

// Enregistrement du fichier Word
$objWriter = IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save($filename);

// Redirection vers une page de succès
header('Location: success.php');
exit();
?>
