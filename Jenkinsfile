pipeline {
    agent any

    environment {
        GITHUB_TOKEN = credentials('github-token')
        SVN_PASS     = credentials('svn-password')
        SVN_USER     = 'svnuser'
        SVN_URL      = 'svn://172.19.0.1:3690/company_repo/Project_Engineering/6_Review/Git_review_log/'
        REPO_LIST    = 'https://github.com/tohuy10/Test-features,https://github.com/tohuy10/Bestarion-Shopping-Application'
    }

    stages {
        stage('Checkout Script Repository') {
            steps {
                // Pulls the repo containing sync_reviews.py directly into the workspace
                git branch: 'main', url: 'https://github.com/tohuy10/Git-to-SVN-IR.git'
            }
        }

        stage('Execute Review Sync') {
            steps {
                sh '''
                    # Verify environment tools
                    python3 --version
                    svn --version | head -n 1

                    # Run the script pulled from the repository
                    python3 sync_reviews.py
                '''
            }
        }
    }

    post {
        success {
            echo "Review logs successfully synced and committed to SVN."
        }
        failure {
            echo "Review sync job failed. Check console output for details."
        }
    }
}