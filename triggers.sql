/*Criação de tabela log de dono_pet para triggers */
CREATE TABLE DONO_PET_LOG ( 
	DONO_PET_ID INT AUTO_INCREMENT PRIMARY KEY, 
	DATA_REG DATE,
	HORA_REG TIME,
	USUARIO VARCHAR (100)
);

/* Trigger 01 - dono_pet_log
Essa trigger ativa depois da inserção de dado no sistema, onde registra a data, hora e nome do usuário
que registrou uma nova inserção de dados
*/
DELIMITER //
CREATE TRIGGER tg_dono_pet_id_after
AFTER INSERT ON DONO_PET
FOR EACH ROW
BEGIN
	INSERT INTO DONO_PET_LOG (DATA_REG, HORA_REG, USUARIO)
		VALUES (NOW(), NOW(), CURRENT_USER());
END;


INSERT INTO DONO_PET 
	(DONO_NOME, DONO_SOBRENOME, DONO_CPF, DONO_ENDERECO, DONO_EMAIL, DONO_TELEFONE)
VALUES
	('Leandro', 'Gois', '45698578596', 'Rua J, 444', 'l.gois@email.com', '457-128-9078');
//

/*Ao selecionar a tabela log, encontraremos data, hora e usuário que inseriu os dados */
SELECT * FROM DONO_PET_LOG;



/* Trigger 02 - dono_pet_log
Essa trigger ativa antes da inserção de dado no sistema, onde manda uma mensagem pro usuário de que
o campo DONO_NOME é obrigatório, não podendo ser em branco '' ou nulo.
*/
DELIMITER //
CREATE TRIGGER tg_dono_pet_id_before
BEFORE INSERT ON DONO_PET
FOR EACH ROW
BEGIN
	IF NEW.DONO_NOME = '' OR NEW.DONO_NOME IS NULL THEN
		SIGNAL SQLSTATE '45000'
		SET MESSAGE_TEXT = 'O campo DONO_NOME não pode ser deixado em branco';
    END IF;
END;
//

/* Ao tentar inserir os dados a seguir, retornará a mensagem de erro da trigger
INSERT INTO DONO_PET 
	(DONO_NOME, DONO_SOBRENOME, DONO_CPF, DONO_ENDERECO, DONO_EMAIL, DONO_TELEFONE)
VALUES
	('', 'Gois', '45698578596', 'Rua J, 444', 'l.gois@email.com', '457-128-9078');
*/




/*Criação de tabela log de consulta para triggers*/
CREATE TABLE CONSULTA_LOG ( 
CONSULTA_ID INT AUTO_INCREMENT PRIMARY KEY, 
DATA_REGISTRO DATE,
HORA_REGISTRO TIME,
USUARIO VARCHAR (100)
);


/* Trigger 01 - consulta_log
Essa trigger ativa depois da inserção de dado no sistema, onde registra a data, hora e nome do usuário
que registrou uma nova inserção de dados
*/
DELIMITER //
CREATE TRIGGER tg_consulta_after
AFTER INSERT ON CONSULTA
FOR EACH ROW
BEGIN
  
INSERT INTO CONSULTA_LOG (DATA_REGISTRO, HORA_REGISTRO, USUARIO)
    VALUES (NOW(), NOW(), CURRENT_USER());
END;


INSERT INTO CONSULTA
VALUES
	(503, 8, 4,9, '2023-10-30', '18:00');
//

/*Ao selecionar a tabela log, encontraremos data, hora e usuário que inseriu os dados */
SELECT * FROM CONSULTA_LOG; 


/* Trigger 02 - consulta_log
Essa trigger ativa antes da inserção de dado no sistema, onde manda uma mensagem pro usuário de que
o campo ID_CONSULTA é obrigatório, não podendo menor que 0.
*/
DELIMITER //
CREATE TRIGGER tg_consulta_before
BEFORE INSERT ON CONSULTA
FOR EACH ROW
BEGIN
    IF NEW.ID_CONSULTA < 0 THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'Não é permitido inserir valores negativos em minha_coluna';
    END IF;
END;
//

/* Se alguém tentar inserir um valor negativo na coluna CONSULTA_ID, por exemplo -1,
 a trigger gerará um erro com a mensagem especificada.
 */









 

