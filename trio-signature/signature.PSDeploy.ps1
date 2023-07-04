Deploy SignatureDeployment {

    By FileSystem Scripts {

        FromSource 'Signatures'
        To '\\filesrv.trio.local\e$\Signatures\'
    }
}
