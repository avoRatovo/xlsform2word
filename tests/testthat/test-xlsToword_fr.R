test_that("xlsform2word_fr retourne 'Success' et génère un fichier Word pour chaque formulaire", {
  # Liste des fichiers Excel à tester
  test_files <- list.files("test_files", pattern = "\\.xlsx?$", full.names = TRUE)

  # Boucle sur chaque fichier
  for (f in test_files) {
    cat("📄 Test du fichier :", basename(f), "\n")

    output_path <- tempfile(fileext = ".docx")

    # Exécution de la fonction
    result <- xlsform2word_fr(f, output = output_path)

    # Vérifications
    expect_equal(result, "Success")          # La fonction retourne "Success"
    expect_true(file.exists(output_path))    # Le fichier Word est bien créé
  }
})
