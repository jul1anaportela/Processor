from setuptools import setup, find_packages

setup(
    name='projeto_processor',
    version='1.0.0',
    packages=find_packages(),
    install_requires=[
        'pandas>=1.0.0',
        'openpyxl>=3.0.0',
        'pytest'
    ],
    entry_points={
        'console_scripts': [
            'processar-participantes=projeto_processor.main:main',
        ]
    },
    author='Juliana Portela',
    description='Processa arquivos Excel de participantes e exporta para JSON por projeto.',
    python_requires='>=3.6',
    include_package_data=True,
    zip_safe=False, 
    license='MIT',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python'
    ]
)