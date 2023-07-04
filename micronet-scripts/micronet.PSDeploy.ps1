Deploy ExampleDeployment {

    By FileSystem Scripts {

        FromSource 'micronet'
        To '\\micronet\c$\scripts\'
    }
}
