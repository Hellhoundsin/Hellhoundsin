# Criado em VBA por Victor Braga em nov/2009
# Traduzido pro R por Victor Braga em mai/2022
# Disponibilizado em erandom.com.br

#### Input ####

##### Libraries ####

require(readxl)
require(dplyr)
require(lubridate)
require(doParallel)
require(matrixStats)
require(foreach)
require(tictoc)
require(tidyr) #para heatmap
require(ggplot2) #para heatmap
require(scales) #para heatmap


##### Variáveis de usuário ####

PastadeTrabalho <- "R:\\Victor\\GoogleDrive\\Composições\\Programas\\Excel, VBA\\Campeonato\\Brasileiro 2022\\"
Scraper <- "ScraperBrasileiro3.xlsm"
Sheet <- "SerieA"

RodarMultithread <- TRUE

Ano <- 2022
Simulacoes <- 100000

PorData <- TRUE
DataInicio <- dmy_hms("09-04-2022 00:00", truncated = 1)
DataFim <- dmy_hms("31-12-2022 23:59:59", truncated = 1)
PorRodada <- !PorData
RodadaInicio <- 0 #not used yet
RodadaFim <- 4 #not used yet
Divisao <- ifelse(Sheet == "SerieA", 1, 2)


##### Variáveis de sistema ####

ModelName <- paste("MCMC Campeonato Brasileiro ", Ano, " by Victor Braga", sep="")

PesoGolFeito <- 0.8
PesoGolLevado <- 1 - PesoGolFeito
FreqBar <- ceiling(sqrt(Simulacoes))

RangeLibertadores <- c(1, 6)
RangeLibertadores2Fase <- c(1, 4)
RangeSulAmericana <- c(7, 12)
RangeRebaixamento <- c(17, 20)
RangeAscensao <- c(1, 4)

options("scipen"=999)
n.cores <- ifelse(RodarMultithread, detectCores() - 1, 1)



#### Importação ####

tic("Run Total")
setwd(PastadeTrabalho)

Jogos <- as.data.frame(read_excel(Scraper, sheet = Sheet,
                                  col_types = c("text", "numeric", "date", "date", "text", "numeric", "numeric", 
                                                "text", "text", "text", "skip", "skip", "skip")))
hour(Jogos$Data) <- hour(Jogos$Hora)
minute(Jogos$Data) <- minute(Jogos$Hora)
second(Jogos$Data) <- second(Jogos$Hora)
Jogos = select(Jogos, -Hora, -Cidade, -Estadio)

Times <- as.data.frame(sort(union(Jogos$Time1, Jogos$Time2)))
colnames(Times) <- c("Time")


#### Análise atual do resultado ####

Jogos$R_V <- Jogos$Gols1 > Jogos$Gols2
Jogos$R_E <- Jogos$Gols1 == Jogos$Gols2
Jogos$R_D <- Jogos$Gols1 < Jogos$Gols2

Lado1 <- Jogos %>%
    filter(!is.na(Gols1)) %>%
    group_by(Time1) %>% 
    summarize(
        Jogos = n(),
        Vitorias = sum(R_V, na.rm = TRUE),
        Empates = sum(R_E, na.rm = TRUE),
        Derrotas = sum(R_D, na.rm = TRUE),
        GolsPro = sum(Gols1, na.rm = TRUE),
        GolsContra = sum(Gols2, na.rm = TRUE))

Lado2 <- Jogos %>%
    filter(!is.na(Gols2)) %>%
    group_by(Time2) %>% 
    summarize(
        Jogos = n(),
        Vitorias = sum(R_D, na.rm = TRUE),
        Empates = sum(R_E, na.rm = TRUE),
        Derrotas = sum(R_V, na.rm = TRUE),
        GolsPro = sum(Gols2, na.rm = TRUE),
        GolsContra = sum(Gols1, na.rm = TRUE))

Times$Jogos <- Lado1$Jogos + Lado2$Jogos
Times$Vitorias <- Lado1$Vitorias + Lado2$Vitorias
Times$Empates <- Lado1$Empates + Lado2$Empates
Times$Derrotas <- Lado1$Derrotas + Lado2$Derrotas
Times$Pontos <- Times$Vitorias * 3 + Times$Empates
Times$Aproveitamento <- ifelse(Times$Jogos > 0, Times$Pontos / (Times$Jogos * 3), 1)
Times$GolsPro <- Lado1$GolsPro + Lado2$GolsPro
Times$GolsContra <- Lado1$GolsContra + Lado2$GolsContra
Times$SaldoGols <- Times$GolsPro - Times$GolsContra

salt <- runif(nrow(Times))
Times$Desempate <- Times$Pontos * 10000000 + Times$Vitorias * 100000 + Times$SaldoGols * 1000 + Times$GolsPro + salt
Times$Posicao <- nrow(Times) - rank(Times$Desempate) + 1
Jogos <- select(Jogos, -R_V, -R_E, -R_D)
Times <- select(Times, -Desempate)



#### Simulações ####

##### Cálculo paramétrico ####

Lambdas <- as.data.frame(Times$Time)
names(Lambdas) <- c("Time")
Lambdas$MediaGolsPro <- ifelse(Times$Jogos > 0, Times$GolsPro / Times$Jogos, 0)
Lambdas$MediaGolsMandantePro <- ifelse(Lado1$Jogos > 0, Lado1$GolsPro / Lado1$Jogos, 1)
Lambdas$MediaGolsVisitantePro <- ifelse(Lado2$Jogos > 0, Lado2$GolsPro / Lado2$Jogos, 1)
Lambdas$MediaGolsContra <- ifelse(Times$Jogos > 0, Times$GolsContra / Times$Jogos, 0)
Lambdas$MediaGolsMandanteContra <- ifelse(Lado1$Jogos > 0, Lado1$GolsContra / Lado1$Jogos, 1)
Lambdas$MediaGolsVisitanteContra <- ifelse(Lado2$Jogos > 0, Lado2$GolsContra / Lado2$Jogos, 1)

JogosSim <- Jogos %>% filter(is.na(Gols1))
JogosSim <- JogosSim %>% filter(Data >= DataInicio)
JogosSim <- JogosSim %>% filter(Data <= DataFim)
JogosASimular <- nrow(JogosSim)

JogosSim <- left_join(JogosSim, Lambdas %>% select(Time, MediaGolsMandantePro, MediaGolsMandanteContra), by=c("Time1" = "Time"))
JogosSim <- left_join(JogosSim, Lambdas %>% select(Time, MediaGolsVisitantePro, MediaGolsVisitanteContra), by=c("Time2" = "Time"))

JogosSim$lambda1 <- PesoGolFeito * JogosSim$MediaGolsMandantePro + PesoGolLevado * JogosSim$MediaGolsVisitanteContra
JogosSim$lambda2 <- PesoGolFeito * JogosSim$MediaGolsVisitantePro + PesoGolLevado * JogosSim$MediaGolsMandanteContra


##### Matriz de simulação ####

Gols1 <- vector(mode="integer", length=JogosASimular * Simulacoes)
dim(Gols1) <- c(JogosASimular, Simulacoes)
Gols2 <- vector(mode="integer", length=JogosASimular * Simulacoes)
dim(Gols2) <- c(JogosASimular, Simulacoes)

for (i in 1:JogosASimular){
    Gols1[i,] <- rpois(Simulacoes, JogosSim[i,"lambda1"])
    Gols2[i,] <- rpois(Simulacoes, JogosSim[i,"lambda2"])
}


##### Criando multithread ####

my.cluster <- makeCluster(n.cores, type = "PSOCK")
print(my.cluster)
registerDoParallel(cl = my.cluster)
getDoParRegistered()
getDoParWorkers()

##### Consolidação Simulações ####

Posicoes <- vector(mode="integer", length=nrow(Times) * Simulacoes)
dim(Posicoes) <- c(nrow(Times), Simulacoes)

Posicoes <- foreach(k = 1:Simulacoes, .combine = 'c', .packages=c("dplyr")) %dopar% {
    
    ###### redefinindo jogos ####
    
    JogosSimK <- select(JogosSim, Key, Gols1, Gols2)
    JogosSimK$Gols1 <- Gols1[, k]
    JogosSimK$Gols2 <- Gols2[, k]
    JogosK <- Jogos
    JogosK <- left_join(JogosK, JogosSimK, by=c("Key" = "Key"))
    JogosK$Gols1 <- ifelse(is.na(JogosK$Gols1.x), JogosK$Gols1.y, JogosK$Gols1.x)
    JogosK$Gols2 <- ifelse(is.na(JogosK$Gols2.x), JogosK$Gols2.y, JogosK$Gols2.x)
    JogosK <- select(JogosK, -Gols1.x, -Gols1.y, -Gols2.x, -Gols2.y)
    
    ##### análise simulação ####
    
    TimesK <- as.data.frame(sort(union(JogosK$Time1, JogosK$Time2)))
    colnames(TimesK) <- c("Time")
    
    JogosK$R_V <- JogosK$Gols1 > JogosK$Gols2
    JogosK$R_E <- JogosK$Gols1 == JogosK$Gols2
    JogosK$R_D <- JogosK$Gols1 < JogosK$Gols2
    
    Lado1 <- JogosK %>%
        filter(!is.na(Gols1)) %>%
        group_by(Time1) %>% 
        summarize(
            Vitorias = sum(R_V, na.rm = TRUE),
            Empates = sum(R_E, na.rm = TRUE),
            GolsPro = sum(Gols1, na.rm = TRUE),
            GolsContra = sum(Gols2, na.rm = TRUE))
    
    Lado2 <- JogosK %>%
        filter(!is.na(Gols2)) %>%
        group_by(Time2) %>% 
        summarize(
            Vitorias = sum(R_D, na.rm = TRUE),
            Empates = sum(R_E, na.rm = TRUE),
            GolsPro = sum(Gols2, na.rm = TRUE),
            GolsContra = sum(Gols1, na.rm = TRUE))
    
    TimesK$Vitorias <- Lado1$Vitorias + Lado2$Vitorias
    TimesK$Empates <- Lado1$Empates + Lado2$Empates
    TimesK$Pontos <- TimesK$Vitorias * 3 + TimesK$Empates
    TimesK$GolsPro <- Lado1$GolsPro + Lado2$GolsPro
    TimesK$GolsContra <- Lado1$GolsContra + Lado2$GolsContra
    TimesK$SaldoGols <- TimesK$GolsPro - TimesK$GolsContra
    salt <- runif(nrow(TimesK))
    TimesK$Desempate <- TimesK$Pontos * 10000000 + TimesK$Vitorias * 100000 + TimesK$SaldoGols * 1000 + TimesK$GolsPro + salt
    TimesK$Posicao <- nrow(TimesK) - rank(TimesK$Desempate) + 1
    Posicoes[, k] <- TimesK$Posicao
}
stopCluster(cl = my.cluster)
dim(Posicoes) <- c(nrow(Times), Simulacoes)

##### matriz de probabilidades ####

ProbPosicao <- vector(mode="numeric", length=nrow(Times) * nrow(Times))
dim(ProbPosicao) <- c(nrow(Times), nrow(Times))
colnames(ProbPosicao) <- c(paste(c(1:20), "°", sep = ""))
for (k in 1:nrow(Times)){
    ProbPosicao[, k] <- rowCounts(Posicoes, value = k)
}
ProbPosicao <- ProbPosicao / Simulacoes

if (Divisao == 1){
    ProbSituacao <- vector(mode="numeric", length=nrow(Times) * 5)
    dim(ProbSituacao) <- c(nrow(Times), 5)
    colnames(ProbSituacao) <- c("Campeão", "Libertadores", "Libertadores 2ª Fase", "SulAmericana", "Rebaixamento")
    
    ProbSituacao[, 1] <- ProbPosicao[, 1]
    for(k in RangeLibertadores[1]:RangeLibertadores[2]){ProbSituacao[, 2] <- ProbSituacao[, 2] + ProbPosicao[, k]}
    for(k in RangeLibertadores2Fase[1]:RangeLibertadores2Fase[2]){ProbSituacao[, 3] <- ProbSituacao[, 3] + ProbPosicao[, k]}
    for(k in RangeSulAmericana[1]:RangeSulAmericana[2]){ProbSituacao[, 4] <- ProbSituacao[, 4] + ProbPosicao[, k]}
    for(k in RangeRebaixamento[1]:RangeRebaixamento[2]){ProbSituacao[, 5] <- ProbSituacao[, 5] + ProbPosicao[, k]}
} else {
    ProbSituacao <- vector(mode="numeric", length=nrow(Times) * 3)
    dim(ProbSituacao) <- c(nrow(Times), 3)
    colnames(ProbSituacao) <- c("Campeão", "Ascensão", "Rebaixamento")
    
    ProbSituacao[, 1] <- ProbPosicao[, 1]
    for(k in RangeAscensao[1]:RangeAscensao[2]){ProbSituacao[, 2] <- ProbSituacao[, 2] + ProbPosicao[, k]}
    for(k in RangeRebaixamento[1]:RangeRebaixamento[2]){ProbSituacao[, 3] <- ProbSituacao[, 3] + ProbPosicao[, k]}
}
toc()
tic.clearlog()



#### Outputs ####

View(Times)
View(cbind(Times[,1], ProbPosicao))
View(cbind(Times[,1], ProbSituacao))



#### Match Heatmap #####

# Fonte: https://towardsdatascience.com/forecasting-football-scores-with-ggplot2-949de7c1cb52

ScoreGrid<-function(homeXg, awayXg){
    A <- as.numeric()
    B <- as.numeric()
    for(i in 0:9) {
        A[(i+1)] <- dpois(i, homeXg)
        B[(i+1)] <- dpois(i, awayXg)
    }
    A[11] <- 1 - sum(A[1:10])
    B[11] <- 1 - sum(B[1:10])
    name <- c("0","1","2","3","4","5","6","7","8","9","10+")
    zero <- mat.or.vec(11, 1)
    C <- data.frame(row.names=name, "0"=zero, "1"=zero, "2"=zero, "3"=zero, "4"=zero,
                    "5"=zero, "6"=zero, "7"=zero,"8"=zero,"9"=zero,"10+"=zero)
    for(j in 1:11) {
        for(k in 1:11) {
            C[j,k] <- A[k]*B[j]
        }
    }
    colnames(C) <- name
    return(round(C*100, 2)/100)
}

ScoreHeatMap <- function(home, away, homeXg, awayXg, datasource){
    adjustedHome <- as.character(sub("_", " ", home))
    adjustedAway <- as.character(sub("_", " ", away))
    df<-ScoreGrid(homeXg, awayXg)
    df %>%
        as_tibble(rownames = all_of(away)) %>%
        pivot_longer(cols = -all_of(away),
                     names_to = home,
                     values_to = "Prob.") %>%
        mutate_at(vars(all_of(away), home),
                  ~forcats::fct_relevel(.x, "10+", after = 10)) %>%
        ggplot() +
        geom_tile(aes_string(x=all_of(away), y=home, fill = "Prob.")) +
        scale_fill_gradient2(mid="white", high = muted("green"))+
        theme(plot.margin = unit(c(1,1,1,1), "cm"),
              plot.title = element_text(size=20, hjust = 0.5, face="bold",vjust =4),
              plot.caption = element_text(hjust=1.1, size=10, face = "italic"),
              plot.subtitle = element_text(size=12, hjust = 0.5, vjust=4),
              axis.title.x=element_text(size=14, vjust=-0.5, face="bold"),
              axis.title.y=element_text(size=14, vjust =0.5, face="bold"))+
        labs(x=adjustedAway,y=adjustedHome,
             caption=paste("Fonte:", datasource))+
        ggtitle(label = "Expectativa de gols", subtitle=paste(adjustedHome, "x", adjustedAway))
}

View(Jogos[,-1])

JogoHM <- as.numeric(readline("Para qual jogo quer produzir o heatmap (inserir nº)? > "))

if(!is.na(JogoHM)){
    HM_T1 <- Jogos[JogoHM, "Time1"]
    HM_T2 <- Jogos[JogoHM, "Time2"]
    
    HM_L1 <- as.numeric(JogosSim %>%
                            filter(Time1 == HM_T1, Time2 == HM_T2) %>%
                            select(lambda1))
    
    HM_L2 <- as.numeric(JogosSim %>%
                            filter(Time1 == HM_T1, Time2 == HM_T2) %>%
                            select(lambda2))
    
    ScoreHeatMap(gsub(" ", "_", HM_T1), gsub(" ", "_", HM_T2), HM_L1, HM_L2, ModelName)
}



#### Runtime Stats ####

# Excel 2010: 332 x 1000 = 38.2s            ~ 8.7k/s | 8.7k/c/s
# R for: 332 x 1000 : 1 = 17.1s             ~ 19.4k/s | 19.4k/c/s
# R foreach: 332 x 1000 : 1 = 20.9s         ~ 15.9k/s | 15.9k/c/s
# R doparallel: 332 x 1000 : 11 = 6.7s      ~ 49.6k/s | 4.5k/c/s
# R doparallel: 332 x 10000 : 11 = 38.2s    ~ 86.9k/s | 7.9k/c/s
# R doparallel: 332 x 100000 : 11 = 357.1s  ~ 93.0k/s | 8.5k/c/s
# R doparallel: 332 x 100000 : 6 = 393.9s   ~ 84.3k/s | 14.0k/c/s
