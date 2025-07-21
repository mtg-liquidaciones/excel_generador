module.exports = {
  apps: [{
    name: 'generador_excel',
    script: 'npm',
    args: 'start',
    env_production: {
      NODE_ENV: 'production',
    }
  }]
};