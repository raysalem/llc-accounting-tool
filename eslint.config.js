// ESLint configuration (flat config). Run with `npm run lint`.
const js = require('@eslint/js');
const globals = require('globals');

module.exports = [
    js.configs.recommended,
    {
        files: ['**/*.js'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'commonjs',
            globals: { ...globals.node },
        },
        rules: {
            'no-unused-vars': ['error', { args: 'none', caughtErrors: 'none' }],
            'no-empty': ['error', { allowEmptyCatch: false }],
            'eqeqeq': ['error', 'always'],
            'no-var': 'error',
            'prefer-const': ['error', { destructuring: 'all' }],
        },
    },
    {
        // The local site's page runs in the browser.
        files: ['lib/site/public/**/*.js'],
        languageOptions: { sourceType: 'script', globals: { ...globals.browser } },
    },
    { ignores: ['node_modules/', 'coverage/'] },
];
