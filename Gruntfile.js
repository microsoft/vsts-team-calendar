var url = require('url');
var hashFiles = require('hash-files');

module.exports = function (grunt) {

  // Caller-provided options  
  var storageContainerName = grunt.option('storageContainerName') || 'default';
  var storageContainerUri = grunt.option('storageContainerUri') || 'https://vsoteamcalendar.blob.core.windows.net/' + storageContainerName;
  var publisher = grunt.option('publisher') || 'msdevlabs';

  grunt.loadNpmTasks('grunt-contrib-clean');
  grunt.loadNpmTasks('grunt-contrib-copy');
  grunt.loadNpmTasks('grunt-typescript');
  grunt.loadNpmTasks('grunt-azure-blob');
    
  // Project configuration
  grunt.initConfig({
    dirs: {
      output: {
        packages: "dist/packages/",
        web: "dist/web"
      }
    },
    clean: {
      dist: ['dist']
    }
  });

  grunt.config('typescript', {
    main: {
      src: ['scripts/**/*.ts'],
      dest: '<%= dirs.output.web %>',
      options: {
        module: 'amd',
        target: 'es5',
        rootDir: '.',
        sourceMap: false,
        declaration: false
      }
    }
  });

  grunt.config('copy', {
    main: {
      files: [
        {
          expand: true,
          src: ['css/**/*', 'sdk/**/*', 'calendar*.html'],
          dest: '<%= dirs.output.web %>'
        },
        {
          expand: true,
          src: ['images/**', 'vss-extension.json'],
          dest: '<%= dirs.output.packages %>'
        }
      ]
    }
  });

  grunt.config('prepManifest', {
    main: {
      manifest: '<%= dirs.output.packages %>' + '/vss-extension.json',
      webOutput: '<%= dirs.output.web %>'
    }
  });

  var webHash = null;
  var getAzureDest = function () {
    return webHash;
  }
    
  // Task to copy static content to Azure Blob Storage
  
  var storageSource = {
    expand: true,
    cwd: '<%= dirs.output.web %>',
    src: ['**/*'],
    dest: getAzureDest
  };

  grunt.config('azure-blob', {
    options: {
      containerName: storageContainerName,
      containerDelete: false,
      gzip: false,
      copySimulation: false
    },
    dist: {
      files: [
        storageSource
      ]
    }
  });
    
  // Register build task
  grunt.registerTask('build', ['clean', 'copy', 'typescript']);
  
  // Register custom "prep manifest" task
  grunt.task.registerMultiTask('prepManifest', 'Prep the manifest by updating base URI and publisher', function () {
    webHash = hashFiles.sync({
      files: [this.data.webOutput + "/**"]
    });

    console.log("hash: " + webHash);
    storageSource.dest = webHash;

    console.log(JSON.stringify(grunt.config('azure-blob'), null, 2));

    var newBaseUri = storageContainerUri + '/' + webHash;
    
    // Update extension manifest with updated base URI
    var manifest = grunt.file.readJSON(this.data.manifest);
    
    // Update base URI
    manifest.baseUri = newBaseUri;
      
    // Update publisher
    manifest.publisher = publisher;

    grunt.file.write(this.data.manifest, JSON.stringify(manifest, null, 4));
  }); 
  
  // Register prep package and deploy tasks
  grunt.registerTask('prepPackage', ['build', 'prepManifest']);
  grunt.registerTask('deploy', ['prepPackage', 'azure-blob']);

  grunt.registerTask('default', ['build']);
};
