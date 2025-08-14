
module.exports = {
  apps: [{
    name: 'generador_excel',
    script: 'src/app.js',
    node_args: '--import ./preload.js',
    cwd: __dirname,
    watch: false,
    env_production: {
      NODE_ENV: 'production',
    }
  }]
};