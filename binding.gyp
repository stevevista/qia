{
  'targets': [
    {
      'target_name': 'ole_bindings',
      'sources': [
        'src/main.cc',
        'src/utils.cc',
        'src/disp.cc',
        'src/dispatch_object.cc',
        'src/dispatch_callback.cc',
        'src/unknown_objects.cc'
      ],
      'include_dirs': [
        '<!(node -e "require(\'nan\')")'
      ],
      'dependencies': [
      ],
      'conditions': [
      ]
    }
  ]
}
