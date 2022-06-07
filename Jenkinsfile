pipeline{
  
  agent any

  stages{ 
    stage("build"){
            steps{
              echo 'building the application'
              echo 'application built - updated one more time'
            }
    }
            
    stage("test"){
            steps{
            echo 'testing the applicaton'
            }
    }
            
    stage("deploy"){
            steps{
            echo 'deploying the application'
            }
    }
  }
}
