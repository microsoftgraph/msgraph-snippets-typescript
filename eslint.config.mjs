// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import globals from 'globals';
import js from '@eslint/js';
import tsParser from '@typescript-eslint/parser';
import eslintTypeScript from '@typescript-eslint/eslint-plugin';
import eslintPrettierRecommended from 'eslint-plugin-prettier/recommended';
import header from 'eslint-plugin-header';
header.rules.header.meta.schema = false;

export default [
  {
    ignores: ['**/out'],
  },
  js.configs.recommended,
  eslintPrettierRecommended,
  {
    files: ['**/**.{ts,mjs}'],

    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.node,
        RequestInit: true,
      },

      parser: tsParser,
      ecmaVersion: 6,
      sourceType: 'module',
    },

    plugins: {
      eslintTypeScript,
      header,
    },

    rules: {
      'no-unused-vars': [
        'error',
        {
          args: 'all',
          argsIgnorePattern: '^_',
          caughtErrors: 'all',
          caughtErrorsIgnorePattern: '^_',
          destructuredArrayIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
      'header/header': [
        'error',
        'line',
        [
          ' Copyright (c) Microsoft Corporation.',
          ' Licensed under the MIT license.',
        ],
      ],
      'prettier/prettier': [
        'error',
        {
          singleQuote: true,
          endOfLine: 'auto',
          printWidth: 80,
        },
      ],
    },
  },
];
