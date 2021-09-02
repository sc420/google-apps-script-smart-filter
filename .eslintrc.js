module.exports = {
  env: {
    browser: true,
    es2021: true,
  },
  extends: [
    'airbnb-base',
  ],
  parserOptions: {
    ecmaVersion: 12,
  },
  plugins: [
    'eslint-plugin-html',
  ],
  rules: {
    'no-unused-vars': 'warn',
  },
};
