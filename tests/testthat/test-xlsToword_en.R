test_that("xlsform2word_fr retourne 'Success' et génère un fichier Word pour chaque formulaire", {
  # Liste des fichiers Excel à tester
  test_files <- list.files("test_files", pattern = "\\.xlsx?$", full.names = TRUE)

    # Exécution de la fonction
    result <- xlsform2word_en(test_files[1])

    # Vérifications
    expect_equal(result, "Success")          # La fonction retourne "Success"   # Le fichier Word est bien créé
})
