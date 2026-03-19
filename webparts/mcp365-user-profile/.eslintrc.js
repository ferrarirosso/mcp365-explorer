require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      'parserOptions': {
        'project': './tsconfig.json',
        'ecmaVersion': 2018,
        'sourceType': 'module'
      },
      rules: {
        '@rushstack/no-new-null': 1,
        '@rushstack/hoist-jest-mock': 1,
        '@rushstack/import-requires-chunk-name': 1,
        '@rushstack/pair-react-dom-render-unmount': 1,
        '@rushstack/security/no-unsafe-regexp': 1,
        '@typescript-eslint/explicit-function-return-type': [
          1,
          {
            'allowExpressions': true,
            'allowTypedFunctionExpressions': true,
            'allowHigherOrderFunctions': false
          }
        ],
        '@typescript-eslint/explicit-member-accessibility': 0,
        '@typescript-eslint/no-explicit-any': 1,
        '@typescript-eslint/no-floating-promises': 2,
        '@typescript-eslint/no-for-in-array': 2,
        '@typescript-eslint/no-misused-new': 2,
        '@typescript-eslint/no-unused-vars': [1, { 'vars': 'all', 'args': 'none' }],
        '@typescript-eslint/no-use-before-define': [2, { 'functions': false, 'classes': true, 'variables': true, 'enums': true, 'typedefs': true }],
        '@typescript-eslint/no-var-requires': 'error',
        '@typescript-eslint/no-inferrable-types': 0,
        '@typescript-eslint/no-empty-interface': 0,
        'eqeqeq': 1,
        'guard-for-in': 2,
        'max-lines': ['warn', { max: 2000 }],
        'no-eval': 1,
        'no-var': 2,
        'prefer-const': 1,
        'use-isnan': 2,
        '@microsoft/spfx/no-require-ensure': 2,
      }
    },
    {
      files: [
        '*.test.ts', '*.test.tsx', '*.spec.ts', '*.spec.tsx',
        '**/__mocks__/*.ts', '**/__mocks__/*.tsx'
      ],
      rules: {}
    }
  ]
};
